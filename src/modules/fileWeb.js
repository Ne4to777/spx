/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	CACHE_RETRIES_LIMIT,
	getInstance,
	methodEmpty,
	prepareResponseJSOM,
	load,
	executorJSOM,
	isStringEmpty,
	getParentUrl,
	prependSlash,
	convertFileContent,
	setFields,
	hasUrlTailSlash,
	mergeSlashes,
	getFolderFromUrl,
	getFilenameFromUrl,
	executorREST,
	prepareResponseREST,
	getWebRelativeUrl,
	switchCase,
	typeOf,
	shiftSlash,
	ifThen,
	isArrayFilled,
	pipe,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	constant,
	webReport,
	removeEmptyFilenames,
	hasUrlFilename,
	isNumberFilled,
	deep1Iterator,
	deep1IteratorREST
} from '../lib/utility'

import * as cache from '../lib/cache'

// Internal

const liftFolderType = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context)
		if (context.Url !== '/') newContext.Url = shiftSlash(newContext.Url)
		if (!context.Url && context.ServerRelativeUrl) newContext.Url = context.ServerRelativeUrl
		if (!context.ServerRelativeUrl && context.Url) newContext.ServerRelativeUrl = context.Url
		return newContext
	},
	string: (contextUrl = '') => {
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
		super(value)
		this.joinProp = 'ServerRelativeUrl'
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([map(liftFolderType), removeEmptyUrls, removeDuplicatedUrls]),
				constant([liftFolderType()])
			])(value)
			: liftFolderType(value)
	}

	getCount() {
		return this.isArray ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0
	}
}

class FileWeb {
	constructor(parent, files) {
		this.name = 'file'
		this.parent = parent
		this.box = getInstance(Box)(files)
		this.contextUrl = parent.box.head().Url
		this.getContextSPObject = parent.getSPObject
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})

		this.iteratorREST = deep1IteratorREST({
			elementBox: this.parent.box
		})
	}

	async	get(opts = {}) {
		if (opts.asBlob) {
			const { contextUrl } = this
			const result = await this.iteratorREST(({ element }) => {
				const elementUrl = getWebRelativeUrl(contextUrl)(element)
				return executorREST(contextUrl)({
					url: `${this.getRESTObject(elementUrl, contextUrl)}/$value`,
					binaryStringResponseBody: true
				})
			})
			return prepareResponseREST(opts)(result)
		}
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			const parentSPObject = this.getContextSPObject(clientContext)
			const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, parentSPObject)
				: this.getSPObject(elementUrl, parentSPObject)
			return load(clientContext)(spObject)(options)
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(options)))

		return prepareResponseJSOM(options)(result)
	}

	async	create(opts = {}) {
		let needToRetry
		let isError
		const cacheUrl = ['fileCreationRetries', this.parent.id]
		if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const { Content = '', Overwrite = true } = element
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			const parentSPObject = this
				.getSPObjectCollection(getParentUrl(elementUrl), this.getContextSPObject(clientContext))
			const fileCreationInfo = new SP.FileCreationInformation()
			setFields({
				set_url: getFilenameFromUrl(elementUrl),
				set_content: convertFileContent(Content),
				set_overwrite: Overwrite
			})(fileCreationInfo)
			const spObject = parentSPObject.add(fileCreationInfo)
			return load(clientContext)(spObject)(options)
		})
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
					isError = true
					if (err.get_message() === 'File Not Found.') {
						const foldersToCreate = {}
						await this.iterator(({ element }) => {
							const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
							foldersToCreate[getFolderFromUrl(elementUrl)] = true
						})
						const res = await this
							.parent
							.folder(Object.keys(foldersToCreate))
							.create({ silentInfo: true, expanded: true, view: ['Name'] })
							.then(() => {
								const cacheRetriesUrl = ['fileCreationRetries', this.parent.id]
								const retries = cache.get(cacheRetriesUrl)
								if (retries) {
									cache.set(retries - 1)(cacheRetriesUrl)
									return true
								}
								return false
							})
							.catch(error => {
								if (/already exists/.test(error.get_message())) needToRetry = true
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
			return prepareResponseJSOM(opts)(result)
		}
		return undefined
	}

	async	update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const { Content } = element
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const binaryInfo = new SP.FileSaveBinaryInformation()
			if (Content !== undefined) binaryInfo.set_content(convertFileContent(Content))
			const spObject = this.getSPObject(elementUrl, this.getContextSPObject(clientContext))
			spObject.saveBinary(binaryInfo)
			return spObject
		})
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext)(opts)
			}
		}
		this.report('update', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	delete(opts = {}) {
		const { noRecycle } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const parentSPObject = this.getContextSPObject(clientContext)
			const spObject = this.getSPObject(elementUrl, parentSPObject)
			methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	copy(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const parentSPObject = this.getContextSPObject(clientContext)
			const spObject = this.getSPObject(elementUrl, parentSPObject)
			spObject.copyTo(getWebRelativeUrl(contextUrl)({ Url: element.To, Folder: element.Folder }))
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('copy', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	move(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const parentSPObject = this.getContextSPObject(clientContext)
			const spObject = this.getSPObject(elementUrl, parentSPObject)
			spObject.moveTo(getWebRelativeUrl(contextUrl)({ Url: element.To, Folder: element.Folder }))
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('move', opts)
		return prepareResponseJSOM(opts)(result)
	}

	getSPObject(elementUrl, spObject) {
		const { contextUrl } = this
		const folder = getFolderFromUrl(elementUrl)
		const filename = getFilenameFromUrl(elementUrl)
		return spObject.getFileByServerRelativeUrl(
			mergeSlashes(`/${folder ? `${contextUrl}/${folder}` : contextUrl}/${filename}`)
		)
	}

	getSPObjectCollection(elementUrl, spObject) {
		const { contextUrl } = this
		const folder = getFolderFromUrl(elementUrl)
		return folder
			? spObject.getFolderByServerRelativeUrl(`/${contextUrl}/${folder}`).get_files()
			: spObject.get_rootFolder().get_files()
	}

	getRESTObject(elementUrl, contextUrl) {
		const filename = getFilenameFromUrl(elementUrl)
		const folder = getFolderFromUrl(elementUrl)
		return mergeSlashes(
			`/_api/web/getfilebyserverrelativeurl('${prependSlash(
				folder ? `${contextUrl}/${folder}` : contextUrl
			)}/${filename}')`
		)
	}

	report(actionType, opts = {}) {
		webReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			contextUrl: this.contextUrl
		})
	}

	of(files) {
		return getInstance(this.constructor)(this.parent, files)
	}
}

export default getInstance(FileWeb)
