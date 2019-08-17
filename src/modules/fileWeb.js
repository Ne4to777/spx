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
	switchCase,
	typeOf,
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


const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const lifter = switchCase(typeOf)({
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
		this.getContextSPObject = parent.getSPObject.bind(parent)
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
				return executorREST(contextUrl, {
					url: `${this.getRESTObject(elementUrl, contextUrl)}/$value`,
					binaryStringResponseBody: true
				})
			})
			return prepareResponseREST(result, opts)
		}
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getWebRelativeUrl(this.contextUrl)(element)
			const parentSPObject = this.getContextSPObject(clientContext)
			const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(elementUrl, parentSPObject)
				: this.getSPObject(elementUrl, parentSPObject)
			return load(clientContext, spObject, options)
		})
		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, options)
	}

	async	create(opts = {}) {
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
				.getSPObjectCollection(getParentUrl(elementUrl), this.getContextSPObject(clientContext))
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
				await executorJSOM(clientContext)
			}
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
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
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
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
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('copy', opts)
		return prepareResponseJSOM(result, opts)
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
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('move', opts)
		return prepareResponseJSOM(result, opts)
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
