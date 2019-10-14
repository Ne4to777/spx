/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
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
	switchType,
	shiftSlash,
	pipe,
	removeEmptyUrls,
	removeDuplicatedUrls,
	webReport,
	removeEmptyFilenames,
	hasUrlFilename,
	deep1Iterator,
	deep1IteratorREST
} from '../lib/utility'

const KEY_PROP = 'Url'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const lifter = switchType({
	object: context => {
		const newContext = Object.assign({}, context)
		if (context[KEY_PROP] !== '/') newContext[KEY_PROP] = shiftSlash(newContext[KEY_PROP])
		if (!context[KEY_PROP] && context.ServerRelativeUrl) newContext[KEY_PROP] = context.ServerRelativeUrl
		if (!context.ServerRelativeUrl && context[KEY_PROP]) newContext.ServerRelativeUrl = context[KEY_PROP]
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

	getCount() {
		return this.isArray ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0
	}
}

class FileWeb {
	constructor(parent, files) {
		this.name = 'file'
		this.parent = parent
		this.box = getInstance(Box)(files)
		this.contextUrl = parent.box.getHeadPropValue()
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})

		this.iteratorREST = deep1IteratorREST({
			elementBox: this.box
		})
	}

	async get(opts = {}) {
		const {
			asBlob,
			asString,
			asItem
		} = opts
		if (asBlob || asString) {
			const { contextUrl } = this
			const result = await this.iteratorREST(({ element }) => {
				const elementUrl = getWebRelativeUrl(contextUrl)(element)
				return executorREST(contextUrl, {
					url: `${this.getRESTObject(elementUrl, contextUrl)}/$value`,
					binaryStringResponseBody: !asString
				})
			})
			return asString ? result.body : prepareResponseREST(result, opts)
		}
		const options = asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, clientContext)
				: this.getSPObject(elementUrl, clientContext)
			return load(clientContext, spObject, options)
		})
		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, options)
	}

	async create(opts = {}) {
		const foldersToCreate = this.box.reduce(acc => el => {
			const { Folder } = el
			const folder = getParentUrl(getWebRelativeUrl(this.contextUrl)(el)) || Folder
			if (folder) acc.push(folder)
			return acc
		})

		if (foldersToCreate.length) {
			await this.parent.folder(foldersToCreate).get({ view: 'ServerRelativeUrl' }).catch(async () => {
				await this.parent.folder(foldersToCreate).create({ silent: true })
			})
		}
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const { Content = '', Overwrite = true } = element
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			const parentSPObject = this
				.getSPObjectCollection(getParentUrl(elementUrl), clientContext)
			const fileCreationInfo = new SP.FileCreationInformation()
			setFields({
				set_url: getFilenameFromUrl(elementUrl),
				set_content: convertFileContent(Content),
				set_overwrite: Overwrite
			})(fileCreationInfo)
			const spObject = parentSPObject.add(fileCreationInfo)
			return load(clientContext, spObject, options)
		})
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				await executorJSOM(clientContexts[i])
			}
		}

		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	async update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const { Content } = element
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const binaryInfo = new SP.FileSaveBinaryInformation()
			if (Content !== undefined) binaryInfo.set_content(convertFileContent(Content))
			const spObject = this.getSPObject(elementUrl, clientContext)
			spObject.saveBinary(binaryInfo)
			return spObject
		})
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext)
			}
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
	}

	async delete(opts = {}) {
		const { noRecycle } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	async copy(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			spObject.copyTo(getWebRelativeUrl(contextUrl)({ [KEY_PROP]: element.To, Folder: element.Folder }))
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('copy', opts)
		return prepareResponseJSOM(result, opts)
	}

	async move(opts) {
		const { contextUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(contextUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			spObject.moveTo(getWebRelativeUrl(contextUrl)({ [KEY_PROP]: element.To, Folder: element.Folder }))
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('move', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(elementUrl, clientContext) {
		const { contextUrl } = this
		const folder = getFolderFromUrl(elementUrl)
		const filename = getFilenameFromUrl(elementUrl)
		return this.parent.getSPObject(clientContext).getFileByServerRelativeUrl(
			mergeSlashes(`/${folder ? `${contextUrl}/${folder}` : contextUrl}/${filename}`)
		)
	}

	getSPObjectCollection(elementUrl, clientContext) {
		const { contextUrl } = this
		const folder = getFolderFromUrl(elementUrl)
		const parentSPObject = this.parent.getSPObject(clientContext)
		return folder
			? parentSPObject.getFolderByServerRelativeUrl(`/${contextUrl}/${folder}`).get_files()
			: parentSPObject.get_rootFolder().get_files()
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
