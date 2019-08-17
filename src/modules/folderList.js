/* eslint class-methods-use-this:0 */
import {
	// REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE,
	REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE,
	REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE,
	AbstractBox,
	getInstance,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstanceEmpty,
	setItem,
	hasUrlTailSlash,
	methodEmpty,
	getSPFolderByUrl,
	pipe,
	popSlash,
	getListRelativeUrl,
	switchCase,
	typeOf,
	removeEmptyUrls,
	removeDuplicatedUrls,
	shiftSlash,
	mergeSlashes,
	hasProp,
	listReport,
	getTitleFromUrl,
	isStrictUrl,
	deep1Iterator,
	getClientContext,
	executeJSOM,
	identity,
	fix,
	isObject,
	isString,
	urlSplit,
	isArray,
	isObjectFilled,
	prop
} from '../lib/utility'
import * as cache from '../lib/cache'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const buildFolderTree = acc => element => {
	let folder
	if (isObject(element)) {
		const { Folder, Columns = {} } = element
		folder = Folder || Columns.Folder
	} else if (isString(element)) {
		folder = element
	}
	const folders = pipe([popSlash, shiftSlash, urlSplit])(folder)
	mutateObjectTreeFromArray(acc)()(folders)
	return acc
}

const mutateObjectTreeFromArray = fix(fR => b => parentName => ([h, ...t]) => {
	const base = b
	const currentName = parentName ? `${parentName}/${h}` : h
	if (t.length) {
		if (!base[currentName]) base[currentName] = {}
		return fR(base[currentName])(currentName)(t)
	}
	if (!base[currentName]) base[currentName] = {}
	return undefined
})

const buildFoldersTree = elements => {
	const foldersTree = {}
	if (isArray(elements)) {
		elements.map(buildFolderTree(foldersTree))
	} else {
		buildFolderTree(foldersTree)(elements)
	}
	return foldersTree
}


const lifter = switchCase(typeOf)({
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
		super(value, lifter, arrayValidator)
		this.joinProp = 'ServerRelativeUrl'
	}
}

class FolderList {
	constructor(parent, folders) {
		this.name = 'folder'
		this.parent = parent
		this.box = getInstance(Box)(folders)
		this.hasColumns = this.box.some(hasProp('Columns'))
		this.contextUrl = parent.parent.box.getHeadPropValue()
		this.listUrl = parent.box.getHeadPropValue()
		this.getContextSPObject = parent.parent.getSPObject.bind(parent.parent)
		this.getListSPObject = parent.getSPObject.bind(parent)
		this.count = this.box.getCount()
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
			return load(clientContext, spObject, options)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, options)
	}

	async	create(opts = {}) {
		const { contextUrl, listUrl } = this
		const { asItem } = opts
		const getRelativeUrl = getListRelativeUrl(contextUrl)(listUrl)
		const options = asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		if (this.hasColumns) {
			await this.cacheColumns()
		}

		const property = this.box.prop

		const foldersMap = this.box.reduce(acc => el => {
			acc[el[property]] = el
			return acc
		}, {})
		const filteredResult = []

		await fix(fR => async b => {
			const names = Reflect.ownKeys(b)
			const promises = []

			for (let i = 0; i < names.length; i += 1) {
				const name = names[i]
				const element = foldersMap[name]
				let elementUrl
				let newElement
				if (element) {
					elementUrl = getRelativeUrl(element)
					newElement = { ...element, Title: getTitleFromUrl(elementUrl) }
					delete newElement.Url
					delete newElement.ServerRelativeUrl
				} else {
					elementUrl = name
					newElement = { Title: getTitleFromUrl(elementUrl) }
				}

				const clientContext = getClientContext(contextUrl)
				if (!isStrictUrl(elementUrl)) return undefined
				const contextSPObject = this.getContextSPObject(clientContext)
				const listSPObject = this.getListSPObject(listUrl, contextSPObject)
				const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
				itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder)
				itemCreationInfo.set_leafName(elementUrl)
				const spObject = setItem(
					cache.get(['columns', contextUrl, listUrl]) || {}
				)(newElement)(listSPObject.addItem(itemCreationInfo))
				promises.push(executeJSOM(clientContext, spObject.get_folder(), options).catch(identity))
			}

			const result = prepareResponseJSOM(
				(await Promise.all(promises)).filter(el => typeOf(el) !== 'error'),
				opts
			)

			result.reduce((acc, el) => {
				if (foldersMap[getRelativeUrl({
					Url: el.ServerRelativeUrl
				})]) acc.push(el)
				return acc
			}, filteredResult)
			await Promise.all(names.map(async name => {
				const o = b[name]
				if (isObjectFilled(o)) await fR(o)
				return undefined
			}))
			return undefined
		})(buildFoldersTree(this.box.getIterable().map(prop(property))))

		listReport('create', {
			...opts,
			name: this.name,
			box: getInstance(Box)(filteredResult),
			listUrl: this.listUrl,
			contextUrl: this.contextUrl
		})

		return filteredResult
	}

	async	update(opts = {}) {
		const { contextUrl, listUrl } = this
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		await this.cacheColumns()
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
			return load(clientContext, spObject.get_folder(), options)
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
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
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
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
			listUrl: this.listUrl,
			contextUrl: this.contextUrl
		})
	}

	async	cacheColumns() {
		const { contextUrl, listUrl } = this
		if (!cache.get(['columns', contextUrl, listUrl])) {
			const columns = await this
				.parent
				.column()
				.get({
					view: ['TypeAsString', 'InternalName', 'Title', 'Sealed'],
					mapBy: 'InternalName'
				})
			cache.set(columns)(['columns', contextUrl, listUrl])
		}
	}

	of(folders) {
		return getInstance(this.constructor)(this.parent, folders)
	}
}

export default getInstance(FolderList)
