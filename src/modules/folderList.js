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
	switchType,
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
	reduce,
	ifThen,
	isNotError,
	isDefined,
	isUndefined
} from '../lib/utility'
import * as cache from '../lib/cache'

const KEY_PROP = 'Url'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const buildFolderTree = (acc = {}) => element => {
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

const mutateObjectTreeFromArray = b => parentName => ([h, ...t]) => {
	const base = b
	const currentName = parentName ? `${parentName}/${h}` : h
	if (t.length) {
		if (!base[currentName]) base[currentName] = {}
		return mutateObjectTreeFromArray(base[currentName])(currentName)(t)
	}
	if (!base[currentName]) base[currentName] = {}
	return base
}

const buildFoldersTree = ifThen(isArray)([reduce(buildFolderTree)(), buildFolderTree()])


const lifter = switchType({
	object: context => {
		const newContext = Object.assign({}, context)
		if (isUndefined(context[KEY_PROP])) newContext[KEY_PROP] = context.ServerRelativeUrl || context.FileRef
		if (isUndefined(context.ServerRelativeUrl)) newContext.ServerRelativeUrl = context[KEY_PROP] || context.FileRef
		return newContext
	},
	string: (contextUrl = '') => {
		const url = contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl))
		return {
			[KEY_PROP]: url,
			ServerRelativeUrl: url
		}
	},
	default: () => ({
		[KEY_PROP]: '',
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
		this.count = this.box.getCount()

		this.iterator = bundleSize => deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box,
			bundleSize
		})
	}

	async get(opts = {}) {
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { contextUrl, listUrl } = this
		const { clientContexts, result } = await this.iterator()(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			const isCollection = hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, clientContext)
				: this.getSPObject(elementUrl, clientContext)
			return load(clientContext, spObject, options)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, options)
	}

	async create(opts = {}) {
		const { contextUrl, listUrl } = this
		const { asItem } = opts
		const getRelativeUrl = getListRelativeUrl(contextUrl)(listUrl)
		const options = asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		if (this.hasColumns) {
			await this.cacheColumns()
		}

		const foldersMap = this.box.reduce(acc => el => {
			acc[getRelativeUrl(el)] = el
			return acc
		}, {})

		const filteredResult = []

		await fix(fR => async b => {
			const names = Object.keys(b)

			const promises = []
			for (let i = 0; i < names.length; i += 1) {
				const name = names[i]
				const element = foldersMap[name]
				let elementUrl
				let newColumns

				if (element) {
					const { Columns } = element
					elementUrl = getRelativeUrl(element)
					newColumns = { ...Columns, Title: getTitleFromUrl(elementUrl) }
				} else {
					elementUrl = name
					newColumns = { Title: getTitleFromUrl(elementUrl) }
				}

				const clientContext = getClientContext(contextUrl)
				if (!isStrictUrl(elementUrl)) return undefined
				const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
				itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder)
				itemCreationInfo.set_leafName(elementUrl)
				const spObject = setItem(
					cache.get(['columns', contextUrl, listUrl]) || {}
				)(newColumns)(this.parent.getSPObject(listUrl, clientContext).addItem(itemCreationInfo))
				promises.push(executeJSOM(clientContext, spObject.get_folder(), options).catch(identity))
			}

			const result = prepareResponseJSOM(
				(await Promise.all(promises)).filter(isNotError),
				opts
			)

			result.reduce((acc, el) => {
				if (foldersMap[getRelativeUrl({
					[KEY_PROP]: el.ServerRelativeUrl || el.FileRef
				})]) acc.push(el)
				return acc
			}, filteredResult)
			return Promise.all(names.map(async name => {
				const o = b[name]
				return isObjectFilled(o) && fR(o)
			}))
		})(buildFoldersTree(this.box.getIterable().map(getRelativeUrl)))

		listReport('create', {
			...opts,
			name: this.name,
			box: getInstance(Box)(filteredResult),
			listUrl: this.listUrl,
			contextUrl: this.contextUrl
		})

		return this.box.isArray() ? filteredResult : filteredResult[0]
	}

	async update(opts = {}) {
		const { contextUrl, listUrl } = this
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		await this.cacheColumns()
		const { clientContexts, result } = await this.iterator(
			REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE
		)(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			const { Columns = {}, WelcomePage } = element
			const spObject = this.getSPObject(elementUrl, clientContext)
			if (isObjectFilled(Columns)) {
				setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, Columns))(
					spObject.get_listItemAllFields()
				)
			}
			if (isDefined(WelcomePage)) {
				spObject.set_welcomePage(WelcomePage)
				spObject.update()
			}
			return load(clientContext, spObject, options)
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)

		return prepareResponseJSOM(result, opts)
	}

	async delete(opts = {}) {
		const { noRecycle } = opts
		const { contextUrl, listUrl } = this
		const { clientContexts, result } = await this.iterator(
			REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE
		)(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(elementUrl, clientContext) {
		let folder
		const parentSPObject = this.parent.getSPObject(this.listUrl, clientContext)
		if (elementUrl) {
			folder = getSPFolderByUrl(elementUrl)(parentSPObject.get_rootFolder())
		} else {
			const rootFolder = parentSPObject.get_rootFolder()
			rootFolder.isRoot = true
			folder = rootFolder
		}
		return folder
	}

	getSPObjectCollection(elementUrl, clientContext) {
		return this.getSPObject(popSlash(elementUrl), clientContext).get_folders()
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

	async cacheColumns() {
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
