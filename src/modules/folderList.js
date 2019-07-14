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
	deep2Iterator,
	deep3Iterator,
	isNumberFilled,
	listReport,
	getTitleFromUrl,
	isStrictUrl
} from '../lib/utility'
import * as cache from '../lib/cache'

import web from './web'

// Internal

const NAME = 'folder'

const getSPObject = elementUrl => ifThen(constant(elementUrl))([
	spObject => getSPFolderByUrl(elementUrl)(spObject.get_rootFolder()),
	spObject => {
		const rootFolder = methodEmpty('get_rootFolder')(spObject)
		rootFolder.isRoot = true
		return rootFolder
	}
])

const getSPObjectCollection = elementUrl => pipe([getSPObject(popSlash(elementUrl)), methodEmpty('get_folders')])

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

export const cacheColumns = contextBox => elementBox => deep2Iterator({ contextBox, elementBox })(
	async ({ contextElement, element }) => {
		const contextUrl = contextElement.Url
		const listUrl = element.Url
		if (!cache.get(['columns', contextUrl, listUrl])) {
			const columns = await web(contextUrl)
				.list(listUrl)
				.column()
				.get({
					view: ['TypeAsString', 'InternalName', 'Title', 'Sealed'],
					groupBy: 'InternalName'
				})
			cache.set(columns)(['columns', contextUrl, listUrl])
		}
	}
)

// Inteface

export default parent => elements => {
	const instance = {
		box: getInstance(Box)(elements),
		parent
	}
	const iterator = bundleSize => deep3Iterator({
		contextBox: instance.parent.parent.box,
		parentBox: instance.parent.box,
		elementBox: instance.box,
		bundleSize
	})
	const report = actionType => (opts = {}) => listReport({
		...opts,
		NAME,
		actionType,
		box: instance.box,
		listBox: instance.parent.box,
		contextBox: instance.parent.parent.box
	})

	return {
		get: async (opts = {}) => {
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			const { clientContexts, result } = await iterator()(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const contextSPObject = instance.parent.parent.getSPObject(clientContext)
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
					const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element)
					const isCollection = hasUrlTailSlash(elementUrl)
					const spObject = isCollection
						? getSPObjectCollection(elementUrl)(listSPObject)
						: getSPObject(elementUrl)(listSPObject)
					return load(clientContext)(spObject)(options)
				}
			)
			await instance.parent.parent.box.chain(
				el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(options)))
			)
			return prepareResponseJSOM(options)(result)
		},

		create: async function create(opts = {}) {
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			await cacheColumns(instance.parent.parent.box)(instance.parent.box)
			const cacheUrl = ['folderCreationRetries', instance.parent.parent.id]
			if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
			const { clientContexts, result } = await iterator(REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const listUrl = parentElement.Url
					const contextUrl = contextElement.Url
					const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
					if (!isStrictUrl(elementUrl)) return undefined
					const contextSPObject = instance.parent.parent.getSPObject(clientContext)
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
					const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
					itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder)
					itemCreationInfo.set_leafName(elementUrl)
					const newElement = { ...element, Title: getTitleFromUrl(elementUrl) }
					const spObject = setItem(
						cache.get(['columns', contextUrl, listUrl])
					)(Object.assign({}, newElement))(listSPObject.addItem(itemCreationInfo))
					return load(clientContext)(spObject.get_folder())(options)
				}
			)
			let needToRetry
			let isError
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(async el => {
					const clientContextByUrl = clientContexts[el.Url]
					for (let i = 0; i < clientContextByUrl.length; i += 1) {
						const clientContext = clientContextByUrl[i]
						await executorJSOM(clientContext)({ ...options, silentErrors: true }).catch(async err => {
							const msg = err.get_message()
							if (/already exists/.test(msg)) return
							isError = true
							if (/This operation can only be performed on a file;/.test(msg)) {
								const foldersToCreate = {}
								await iterator(REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE)(
									({ contextElement, parentElement, element }) => {
										const elementUrl = getListRelativeUrl(
											contextElement.Url
										)(parentElement.Url)(element)
										foldersToCreate[getParentUrl(elementUrl)] = true
									}
								)
								await deep2Iterator({
									contextBox: instance.parent.parent.box,
									elementBox: instance.parent.box,
									bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
								})(async ({ contextElement, element }) => {
									const res = await web(contextElement.Url)
										.list(element.Url)
										.folder(Object.keys(foldersToCreate))
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
								})
							} else {
								throw err
							}
						})
						if (needToRetry) break
					}
				})
			}
			if (needToRetry) {
				return create(options)
			}
			if (!isError) {
				report('create')(options)
				return prepareResponseJSOM(opts)(result)
			}
			return undefined
		},

		update: async (opts = {}) => {
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			await cacheColumns(instance.parent.parent.box)(instance.parent.box)
			const { clientContexts, result } = await iterator(REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const contextUrl = contextElement.Url
					const listUrl = parentElement.Url
					const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
					if (!isStrictUrl(elementUrl)) return undefined
					const contextSPObject = instance.parent.parent.getSPObject(clientContext)
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
					const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, element))(
						getSPObject(elementUrl)(listSPObject).get_listItemAllFields()
					)
					return load(clientContext)(spObject.get_folder())(options)
				}
			)
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(
					async el => Promise.all(
						clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))
					)
				)
			}
			report('update')(opts)
			return prepareResponseJSOM(opts)(result)
		},

		delete: async (opts = {}) => {
			const { noRecycle } = opts
			const { clientContexts, result } = await iterator(REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const contextSPObject = instance.parent.parent.getSPObject(clientContext)
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
					const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element)
					if (!isStrictUrl(elementUrl)) return undefined
					const spObject = getSPObject(elementUrl)(listSPObject)
					if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
					return elementUrl
				}
			)
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(
					el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
				)
			}
			report(noRecycle ? 'delete' : 'recycle')(opts)
			return prepareResponseJSOM(opts)(result)
		}
	}
}
