/* eslint class-methods-use-this:0 */
import {
	REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE,
	REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE,
	REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE,
	CACHE_RETRIES_LIMIT,
	AbstractBox,
	getInstance,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstanceEmpty,
	setItem,
	hasUrlTailSlash,
	ifThen,
	constant,
	methodEmpty,
	getSPFolderByUrl,
	pipe,
	popSlash,
	getListRelativeUrl,
	switchCase,
	typeOf,
	isArrayFilled,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	shiftSlash,
	mergeSlashes,
	getParentUrl,
	isNumberFilled,
	listReport,
	getTitleFromUrl,
	isStrictUrl,
	deep1Iterator
} from '../lib/utility'
import * as cache from '../lib/cache'

const liftFolderType = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context)
		if (!context.Url) newContext.Url = context.ServerRelativeUrl || context.FileRef
		if (!context.ServerRelativeUrl) newContext.ServerRelativeUrl = context.Url || context.FileRef
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
}

class FolderList {
	constructor(parent, folders) {
		this.name = 'folder'
		this.parent = parent
		this.box = getInstance(Box)(folders)
		this.contextUrl = parent.contextUrl
		this.getContextSPObject = parent.getContextSPObject
		this.getListSPObject = parent.getListSPObject
		this.iterator = bundleSize => deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box,
			bundleSize
		})
	}

	async	get(opts = {}) {
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { contextUrl, listUrl } = this
		const { clientContexts, result } = await this.iterator()(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			const isCollection = hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, listSPObject)
				: this.getSPObject(elementUrl, listSPObject)
			return load(clientContext)(spObject)(options)
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(options)))
		return prepareResponseJSOM(options)(result)
	}

	async	create(opts = {}) {
		const { contextUrl, listUrl } = this
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		if (!cache.get(['columns', contextUrl, listUrl])) {
			const columns = await this
				.parent
				.column()
				.get({
					view: ['TypeAsString', 'InternalName', 'Title', 'Sealed'],
					groupBy: 'InternalName'
				})
			cache.set(columns)(['columns', contextUrl, listUrl])
		}
		const cacheUrl = ['folderCreationRetries', this.parent.parent.id]
		if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
		const { clientContexts, result } = await this.iterator(
			REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
		)(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
			itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder)
			itemCreationInfo.set_leafName(elementUrl)
			const newElement = { ...element, Title: getTitleFromUrl(elementUrl) }
			const spObject = setItem(
				cache.get(['columns', contextUrl, listUrl])
			)(Object.assign({}, newElement))(listSPObject.addItem(itemCreationInfo))
			return load(clientContext)(spObject.get_folder())(options)
		})
		let needToRetry
		let isError
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext)({ ...options, silentErrors: true }).catch(async err => {
					const msg = err.get_message()
					if (/already exists/.test(msg)) return
					isError = true
					if (/This operation can only be performed on a file;/.test(msg)) {
						const foldersToCreate = {}
						await this.iterator(REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE)(({ element }) => {
							const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
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
								if (/already exists/.test(error.get_message())) return true
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
			return this.create(options)
		}
		if (!isError) {
			this.report('create', options)
			return prepareResponseJSOM(opts)(result)
		}
		return undefined
	}

	async	update(opts = {}) {
		const { contextUrl, listUrl } = this
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		if (!cache.get(['columns', contextUrl, listUrl])) {
			const columns = await this
				.parent
				.column()
				.get({
					view: ['TypeAsString', 'InternalName', 'Title', 'Sealed'],
					groupBy: 'InternalName'
				})
			cache.set(columns)(['columns', contextUrl, listUrl])
		}
		const { clientContexts, result } = await this.iterator(
			REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE
		)(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, element))(
				this.getSPObject(elementUrl, listSPObject).get_listItemAllFields()
			)
			return load(clientContext)(spObject.get_folder())(options)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('update', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	delete(opts = {}) {
		const { noRecycle } = opts
		const { contextUrl, listUrl } = this
		const { clientContexts, result } = await this.iterator(
			REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE
		)(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, listSPObject)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.box.getCount()) {
			await this.parent.parent.box.chain(
				el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
			)
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(opts)(result)
	}

	getSPObject(elementUrl, parentSPObject) {
		let folder
		if (elementUrl) {
			folder = getSPFolderByUrl(elementUrl)(parentSPObject.get_rootFolder())
		} else {
			const rootFolder = parentSPObject.get_rootFolder()
			rootFolder.isRoot = true
			folder = rootFolder
		}
		return folder
	}

	getSPObjectCollection(elementUrl, parentSPObject) {
		return this.getSPObject(popSlash(elementUrl), parentSPObject).get_folders()
	}

	report(actionType, opts = {}) {
		listReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			listBox: this.parent.box,
			contextBox: this.parent.parent.box
		})
	}

	of(folders) {
		return getInstance(this.constructor)(this.parent, folders)
	}
}

export default getInstance(FolderList)
