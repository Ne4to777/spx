/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	getInstance,
	methodEmpty,
	prepareResponseJSOM,
	load,
	executorJSOM,
	switchType,
	getTitleFromUrl,
	popSlash,
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
	deep1Iterator,
	fix,
	getClientContext,
	getParentUrl,
	executeJSOM,
	identity,
	isObjectFilled,
	prop,
	isObject,
	isString,
	urlSplit,
	isArray,
	isDefined,
	isUndefined
} from '../lib/utility'

const KEY_PROP = 'Url'

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

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])
const lifter = switchType({
	object: context => {
		const newContext = Object.assign({}, context)
		if (isUndefined(context[KEY_PROP])) newContext[KEY_PROP] = context.ServerRelativeUrl
		if (context[KEY_PROP] !== '/') newContext[KEY_PROP] = shiftSlash(newContext[KEY_PROP])
		if (!context.ServerRelativeUrl && context[KEY_PROP]) newContext.ServerRelativeUrl = context[KEY_PROP]
		return newContext
	},
	string: contextUrl => {
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

class FolderWeb {
	constructor(parent, folders) {
		this.name = 'folder'
		this.parent = parent
		this.box = getInstance(Box)(folders)
		this.contextUrl = parent.box.getHeadPropValue()
		this.count = this.box.getCount()
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})
	}

	async get(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			const isCollection = hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, clientContext)
				: this.getSPObject(elementUrl, clientContext)
			return load(clientContext, spObject, opts)
		})
		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}

	async create(opts = {}) {
		const { contextUrl } = this
		const getRelativeUrl = getWebRelativeUrl(contextUrl)

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
				const clientContext = getClientContext(contextUrl)
				let elementUrl
				if (element) {
					elementUrl = getRelativeUrl(element)
				} else {
					elementUrl = name
				}

				const parentFolderUrl = getParentUrl(elementUrl)
				const spObject = this
					.getSPObjectCollection(`${parentFolderUrl}/`, clientContext)
					.add(getTitleFromUrl(elementUrl))

				load(clientContext, spObject, opts)

				promises.push(executeJSOM(clientContext, spObject, opts).catch(identity))
			}

			const result = prepareResponseJSOM(
				(await Promise.all(promises)).filter(el => typeOf(el) !== 'error'),
				opts
			)

			result.reduce((acc, el) => {
				if (foldersMap[getRelativeUrl({
					[KEY_PROP]: el.ServerRelativeUrl
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

		webReport('create', {
			...opts,
			name: this.name,
			box: getInstance(Box)(filteredResult),
			contextUrl: this.contextUrl
		})

		return this.box.isArray() ? filteredResult : filteredResult[0]
	}

	async update(opts = {}) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			const { WelcomePage } = element
			const spObject = this.getSPObject(elementUrl, clientContext)
			if (isDefined(WelcomePage)) {
				spObject.set_welcomePage(WelcomePage)
				spObject.update()
			}
			return load(clientContext, spObject, opts)
		})

		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)

		return prepareResponseJSOM(result, opts)
	}

	async delete(opts = {}) {
		const { noRecycle } = opts
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!isStrictUrl(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(elementUrl, clientContext) {
		let folder
		const parentSPObject = this.parent.getSPObject(clientContext)
		if (elementUrl) {
			folder = parentSPObject.getFolderByServerRelativeUrl(elementUrl)
		} else {
			const rootFolder = parentSPObject.get_rootFolder()
			rootFolder.isRoot = true
			folder = rootFolder
		}
		return folder
	}

	getSPObjectCollection(elementUrl, clientContext) {
		return !elementUrl || elementUrl === '/'
			? this.getSPObject(undefined, clientContext).get_folders()
			: this.getSPObject(popSlash(elementUrl), clientContext).get_folders()
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
