/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	CACHE_RETRIES_LIMIT,
	getInstance,
	methodEmpty,
	prepareResponseJSOM,
	load,
	executorJSOM,
	switchCase,
	getTitleFromUrl,
	popSlash,
	getParentUrl,
	pipe,
	hasUrlTailSlash,
	getWebRelativeUrl,
	typeOf,
	shiftSlash,
	mergeSlashes,
	removeEmptyUrls,
	removeDuplicatedUrls,
	webReport,
	isStrictUrl,
	isNumberFilled,
	deep1Iterator
} from '../lib/utility'
import * as cache from '../lib/cache'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])
const lifter = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context)
		if (!context.Url && context.ServerRelativeUrl) newContext.Url = context.ServerRelativeUrl
		if (context.Url !== '/') newContext.Url = shiftSlash(newContext.Url)
		if (!context.ServerRelativeUrl && context.Url) newContext.ServerRelativeUrl = context.Url
		return newContext
	},
	string: contextUrl => {
		const url = contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl))
		return {
			Url: url,
			ServerRelativeUrl: url
		}
	},
	default: () => ({
		Url: '',
		ServerRelativeUrl: ''
	})
})

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter, arrayValidator)
		this.joinProp = 'ServerRelativeUrl'
	}
}

class FolderWeb {
	constructor(parent, folders) {
		this.name = 'folder'
		this.parent = parent
		this.box = getInstance(Box)(folders)
		this.contextUrl = parent.box.getHeadPropValue()
		this.getContextSPObject = parent.getSPObject.bind(parent)
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})
	}

	async	get(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const parentSPObject = this.getContextSPObject(clientContext)
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			const isCollection = hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, parentSPObject)
				: this.getSPObject(elementUrl, parentSPObject)
			return load(clientContext, spObject, opts)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, opts)
	}

	async	create(opts = {}) {
		let needToRetry
		let isError
		const { contextUrl } = this
		const cacheUrl = ['folderCreationRetries', this.parent.id]
		if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const parentFolderUrl = getParentUrl(elementUrl)
			const spObject = this.getSPObjectCollection(`${parentFolderUrl}/`, this.getContextSPObject(clientContext))
				.add(getTitleFromUrl(elementUrl))
			return load(clientContext, spObject, opts)
		})

		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext).catch(async err => {
					const msg = err.message
					if (/already exists/.test(msg)) return
					isError = true
					if (msg === 'File Not Found.') {
						const foldersToCreate = {}
						await this.iterator(({ element }) => {
							const elementUrl = getWebRelativeUrl(contextUrl)(element)
							foldersToCreate[getParentUrl(elementUrl)] = true
						})
						const res = await this
							.of(Object.keys(foldersToCreate))
							.create({ silentInfo: true, expanded: true, view: ['Name'] })
							.then(() => {
								const retries = cache.get(cacheUrl)
								if (retries) {
									cache.set(retries - 1)(cacheUrl)
									return true
								}
								return false
							})
							.catch(error => {
								if (/already exists/.test(error.message)) return true
								return false
							})
						if (res) needToRetry = true
					} else {
						throw err
					}
				})
				if (needToRetry) break
			}
		}
		if (needToRetry) {
			return this.create(opts)
		}
		if (!isError) {
			this.report('create', opts)
			return prepareResponseJSOM(result, opts)
		}
		return undefined
	}

	async	delete(opts = {}) {
		const { noRecycle } = opts
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const parentSPObject = this.getContextSPObject(clientContext)
			const spObject = this.getSPObject(elementUrl, parentSPObject)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(elementUrl, parentSPObject) {
		let folder
		if (elementUrl) {
			folder = parentSPObject.getFolderByServerRelativeUrl(elementUrl)
		} else {
			const rootFolder = parentSPObject.get_rootFolder()
			rootFolder.isRoot = true
			folder = rootFolder
		}
		return folder
	}

	getSPObjectCollection(elementUrl, parentSPObject) {
		return !elementUrl || elementUrl === '/'
			? this.getSPObject(undefined, parentSPObject).get_folders()
			: this.getSPObject(popSlash(elementUrl), parentSPObject).get_folders()
	}


	report(actionType, opts = {}) {
		webReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			contextUrl: this.contextUrl
		})
	}

	of(folders) {
		return getInstance(this.constructor)(this.parent, folders)
	}
}

export default getInstance(FolderWeb)
