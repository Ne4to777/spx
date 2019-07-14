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
	method,
	ifThen,
	constant,
	pipe,
	hasUrlTailSlash,
	getWebRelativeUrl,
	typeOf,
	shiftSlash,
	mergeSlashes,
	isArrayFilled,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	deep2Iterator,
	webReport,
	isStrictUrl,
	isNumberFilled
} from '../lib/utility'
import web from './web'
import * as cache from '../lib/cache'

// Internal

const NAME = 'folder'

const getSPObject = elementUrl => ifThen(constant(elementUrl))([
	method('getFolderByServerRelativeUrl')(elementUrl),
	spObject => {
		const rootFolder = methodEmpty('get_rootFolder')(spObject)
		rootFolder.isRoot = true
		return rootFolder
	}
])

const getSPObjectCollection = elementUrl => pipe([
	ifThen(constant(!elementUrl || elementUrl === '/'))([getSPObject(), getSPObject(popSlash(elementUrl))]),
	methodEmpty('get_folders')
])

const liftFolderType = switchCase(typeOf)({
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

// Inteface

export default parent => elements => {
	const instance = {
		box: getInstance(Box)(elements),
		parent
	}
	const iterator = deep2Iterator({
		contextBox: instance.parent.box,
		elementBox: instance.box
	})

	const report = actionType => (opts = {}) => webReport({
		...opts, NAME, actionType, box: instance.box, contextBox: instance.parent.box
	})
	return {
		get: async opts => {
			const { clientContexts, result } = await iterator(({ contextElement, clientContext, element }) => {
				const parentSPObject = instance.parent.getSPObject(clientContext)
				const contextUrl = contextElement.Url
				const elementUrl = getWebRelativeUrl(contextUrl)(element)
				const isCollection = hasUrlTailSlash(elementUrl)
				const spObject = isCollection
					? getSPObjectCollection(elementUrl)(parentSPObject)
					: getSPObject(elementUrl)(parentSPObject)
				return load(clientContext)(spObject)(opts)
			})
			await instance.parent.box.chain(
				el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
			)
			return prepareResponseJSOM(opts)(result)
		},

		create: async function create(opts = {}) {
			let needToRetry; let
				isError
			const cacheUrl = ['folderCreationRetries', instance.parent.id]
			if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
			const { clientContexts, result } = await iterator(({ contextElement, clientContext, element }) => {
				const elementUrl = getWebRelativeUrl(contextElement.Url)(element)
				if (!isStrictUrl(elementUrl)) return undefined
				const parentFolderUrl = getParentUrl(elementUrl)
				const spObject = getSPObjectCollection(
					`${parentFolderUrl}/`
				)(instance.parent.getSPObject(clientContext)).add(getTitleFromUrl(elementUrl))
				return load(clientContext)(spObject)(opts)
			})

			if (instance.box.getCount()) {
				await instance.parent.box.chain(async el => {
					const clientContextByUrl = clientContexts[el.URL]
					for (let i = 0; i < clientContextByUrl.length; i += 1) {
						const clientContext = clientContextByUrl[i]
						await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
							const msg = err.get_message()
							if (/already exists/.test(msg)) return
							isError = true
							if (msg === 'File Not Found.') {
								const foldersToCreate = {}
								await iterator(({ contextElement, element }) => {
									const elementUrl = getWebRelativeUrl(contextElement.Url)(element)
									foldersToCreate[getParentUrl(elementUrl)] = true
								})
								const res = await web(clientContext.get_url())
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
							} else {
								throw err
							}
						})
						if (needToRetry) break
					}
				})
			}
			if (needToRetry) {
				return create(opts)
			}
			if (!isError) {
				report('create')(opts)
				return prepareResponseJSOM(opts)(result)
			}
			return undefined
		},

		delete: async (opts = {}) => {
			const { noRecycle } = opts
			const { clientContexts, result } = await iterator(({ contextElement, clientContext, element }) => {
				const contextUrl = contextElement.Url
				const elementUrl = getWebRelativeUrl(contextUrl)(element)
				if (!isStrictUrl(elementUrl)) return undefined
				const parentSPObject = instance.parent.getSPObject(clientContext)
				const spObject = getSPObject(elementUrl)(parentSPObject)
				if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
				return elementUrl
			})
			if (instance.box.getCount()) {
				await instance.parent.box.chain(
					el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
				)
			}
			report(noRecycle ? 'delete' : 'recycle')(opts)
			return prepareResponseJSOM(opts)(result)
		}
	}
}
