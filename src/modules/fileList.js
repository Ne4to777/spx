/* eslint class-methods-use-this:0 */
import axios from 'axios'
import {
	ACTION_TYPES,
	LIBRARY_STANDART_COLUMN_NAMES,
	AbstractBox,
	getInstance,
	methodEmpty,
	prepareResponseJSOM,
	getClientContext,
	load,
	executorJSOM,
	convertFileContent,
	setFields,
	hasUrlTailSlash,
	isArray,
	mergeSlashes,
	getFolderFromUrl,
	getFilenameFromUrl,
	executorREST,
	prepareResponseREST,
	isExists,
	getSPFolderByUrl,
	popSlash,
	identity,
	getInstanceEmpty,
	executeJSOM,
	isObject,
	setItem,
	ifThen,
	join,
	getListRelativeUrl,
	listReport,
	switchType,
	shiftSlash,
	pipe,
	removeEmptyUrls,
	removeDuplicatedUrls,
	isUndefined,
	isBlob,
	hasUrlFilename,
	removeEmptyFilenames,
	isObjectFilled,
	deep1Iterator,
	deep1IteratorREST,
	getRequestDigest,
	getParentUrl,
	getWebRelativeUrl
} from '../lib/utility'
import * as cache from '../lib/cache'

const KEY_PROP = 'Url'

async function copyOrMove(isMove, opts = {}) {
	const { contextUrl, listUrl } = this
	await this.iteratorREST(async ({ element }) => {
		const { To, OnlyContent } = element
		const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
		if (!hasUrlFilename(elementUrl)) return
		let targetWebUrl
		let targetListUrl
		let targetFileUrl
		if (isObject(To)) {
			targetWebUrl = To.WebUrl || contextUrl
			targetListUrl = To.ListUrl
			targetFileUrl = getListRelativeUrl(To.WebUrl)(To.ListUrl)(To) || ''
		} else {
			targetWebUrl = contextUrl
			targetListUrl = listUrl
			targetFileUrl = To
		}

		if (!targetWebUrl) throw new Error('Target WebUrl is missed')
		if (!targetListUrl) throw new Error('Target ListUrl is missed')
		if (!elementUrl) throw new Error('Source file Url is missed')
		const spxSourceList = this.parent.of(listUrl)
		const spxSourceFile = spxSourceList.file(elementUrl)
		const spxTargetList = this.parent.parent.of(targetWebUrl).library(targetListUrl)
		const fullTargetFileUrl = hasUrlFilename(targetFileUrl)
			? targetFileUrl
			: `${targetFileUrl}/${getFilenameFromUrl(elementUrl)}`
		const columnsToUpdate = {}
		const existedColumnsToUpdate = {}
		if (!OnlyContent) {
			const sourceFileData = await spxSourceFile.get({ asItem: true })
			const keys = Reflect.ownKeys(sourceFileData)
			for (let i = 0; i < keys.length; i += 1) {
				const columnName = keys[i]
				if (!LIBRARY_STANDART_COLUMN_NAMES[columnName] && sourceFileData[columnName] !== null) {
					columnsToUpdate[columnName] = sourceFileData[columnName]
				}
			}
			if (Object.keys(columnsToUpdate).length) {
				const columnKeys = Reflect.ownKeys(columnsToUpdate)
				for (let i = 0; i < columnKeys.length; i += 1) {
					const columnName = columnKeys[i]
					existedColumnsToUpdate[columnName] = sourceFileData[columnName]
				}
			}
		}
		if (!opts.forced && contextUrl === targetWebUrl) {
			const clientContext = getClientContext(contextUrl)
			const spObject = this.getSPObject(elementUrl, clientContext)
			const folder = getFolderFromUrl(targetFileUrl)
			if (folder) {
				await this
					.parent
					.folder(folder)
					.create({ silentInfo: true, expanded: true, view: ['Name'] })
					.catch(identity)
			}
			spObject[isMove ? 'moveTo' : 'copyTo'](mergeSlashes(`${targetListUrl}/${fullTargetFileUrl}`))
			await executeJSOM(clientContext, spObject, opts)
			await spxTargetList
				.file({ [KEY_PROP]: targetFileUrl, Columns: existedColumnsToUpdate })
				.update({ silentInfo: true })
		} else {
			await spxTargetList
				.file({
					[KEY_PROP]: fullTargetFileUrl,
					Content: await spxSourceList.file(elementUrl).get({ asBlob: true }),
					OnProgress: element.OnProgress,
					Overwrite: element.Overwrite,
					Columns: existedColumnsToUpdate
				})
				.create({ silentInfo: true })
			if (isMove) await spxSourceFile.delete()
		}
	})

	console.log(`${ACTION_TYPES[isMove ? 'move' : 'copy']} ${this.count} ${this.name}(s)`)
}

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const lifter = switchType({
	object: context => {
		const newContext = Object.assign({}, context)
		const name = context.Content ? context.Content.name : undefined
		if (!context[KEY_PROP]) newContext[KEY_PROP] = context.ServerRelativeUrl || context.FileRef || name
		if (!context.ServerRelativeUrl) newContext.ServerRelativeUrl = context[KEY_PROP] || context.FileRef
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

async function createWithRESTFromString(element, opts = {}) {
	const { needResponse } = opts
	const { Content = '', Overwrite = true, Columns } = element
	const { contextUrl, listUrl } = this
	const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
	const filename = getFilenameFromUrl(elementUrl)
	const filesUrl = this.getRESTObjectCollection(elementUrl, listUrl, contextUrl)
	await axios({
		url: `${filesUrl}/add(url='${filename}',overwrite=${Overwrite})`,
		headers: {
			Accept: 'application/json;odata=verbose',
			'Content-Type': 'application/json;odata=verbose',
			'X-RequestDigest': await getRequestDigest()
		},
		method: 'POST',
		data: Content
	})

	let response = element

	if (Columns) {
		response = await this
			.of({ [KEY_PROP]: elementUrl, Columns })
			.update({ ...opts, silentInfo: true })
	} else if (needResponse) {
		response = await this
			.of(elementUrl)
			.get(opts)
	}

	return response
}

async function createWithRESTFromBlob(element, opts = {}) {
	const { needResponse, silent, silentErrors } = opts
	const {
		Content = '',
		Overwrite = true,
		OnProgress = identity,
		Folder = '',
		Columns
	} = element
	const { contextUrl, listUrl } = this

	const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
	const folder = Folder || getFolderFromUrl(elementUrl)
	const filename = elementUrl ? getFilenameFromUrl(elementUrl) : Content.name
	const requiredInputs = {
		__REQUESTDIGEST: true,
		__VIEWSTATE: true,
		__EVENTVALIDATION: true
	}

	const inputs = []
	const listGUID = cache.get(['listGUIDs', contextUrl, listUrl])
	const listFormMatches = cache.get(['listFormMatches', contextUrl, listUrl])
	const inputRE = /<input[^<]*\/>/g
	let founds = inputRE.exec(listFormMatches)

	while (founds) {
		const item = founds[0]
		const id = item.match(/id="([^"]+)"/)[1]
		if (requiredInputs[id]) {
			inputs.push(item)
		}
		founds = inputRE.exec(listFormMatches)
	}

	const form = window.document.createElement('form')
	form.innerHTML = join('')(inputs)
	const formData = new FormData(form)
	formData.append('__EVENTTARGET', 'ctl00$PlaceHolderMain$ctl03$RptControls$btnOK')
	formData.append('ctl00$PlaceHolderMain$ctl04$ctl01$uploadLocation', `/${folder.replace(/^\//, '')}`)
	if (Overwrite) formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$OverwriteSingle', true)
	formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', Content, filename)

	const response = await axios({
		url: `/${contextUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
		method: 'POST',
		data: formData,
		onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
	})

	const errorMsgMatches = response.data.match(/id="ctl00_PlaceHolderMain_LabelMessage">([^<]*)<\/span>/)
	let res = { [KEY_PROP]: elementUrl }
	if (isArray(errorMsgMatches)) {
		if (!silent && !silentErrors) console.error(errorMsgMatches[1])
	} else if (isObjectFilled(Columns)) {
		res = await this.of({ [KEY_PROP]: elementUrl, Columns }).update({ ...opts, silentInfo: true })
	} else if (needResponse) {
		res = await this.of({ [KEY_PROP]: elementUrl }).get(opts)
	}
	return res
}

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter, arrayValidator)
		this.joinProp = 'ServerRelativeUrl'
	}

	getCount() {
		return this.isArray() ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0
	}
}

class FileList {
	constructor(parent, files) {
		this.name = 'file'
		this.parent = parent
		this.box = getInstance(Box)(files)
		this.count = this.box.getCount()
		this.contextUrl = parent.parent.box.getHeadPropValue()
		this.listUrl = parent.box.getHeadPropValue()
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})

		this.iteratorREST = deep1IteratorREST({
			elementBox: this.box
		})
	}

	async get(opts = {}) {
		const { contextUrl, listUrl } = this
		if (opts.asBlob) {
			const result = await this.iteratorREST(({ element }) => {
				const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
				return executorREST(contextUrl, {
					url: `${this.getRESTObject(elementUrl, listUrl, contextUrl)}/$value`,
					binaryStringResponseBody: true
				})
			})
			return prepareResponseREST(result, opts)
		}
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			const spObject = isExists(elementUrl) && hasUrlTailSlash(elementUrl)
				? this.getSPObjectCollection(elementUrl, clientContext)
				: this.getSPObject(elementUrl, clientContext)
			return load(clientContext, spObject, options)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, options)
	}

	async create(opts = {}) {
		const { contextUrl, listUrl } = this
		if (!cache.get(['listGUIDs', contextUrl, listUrl])) {
			const listProps = await this
				.parent
				.of(listUrl)
				.get({ view: 'Id' })
			cache.set(listProps.Id.toString())(['listGUIDs', contextUrl, listUrl])
		}
		if (!cache.get(['listFormMatches', contextUrl, listUrl])) {
			const listForms = await axios.get(
				`/${contextUrl}/_layouts/15/Upload.aspx?List={${cache.get(['listGUIDs', contextUrl, listUrl])}}`
			)
			cache.set(listForms.data.match(/<form(\w|\W)*<\/form>/))(['listFormMatches', contextUrl, listUrl])
		}

		const foldersToCreate = this.box.reduce(acc => el => {
			const { Folder } = el
			const folder = getParentUrl(getWebRelativeUrl(contextUrl)(el)) || Folder
			if (folder) acc.push(folder)
			return acc
		})

		if (foldersToCreate.length) {
			await this.parent.folder(foldersToCreate).get({ view: 'ServerRelativeUrl' }).catch(async () => {
				await this.parent.folder(foldersToCreate).create({ silent: true })
			})
		}

		const res = await this.iteratorREST(({ element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!hasUrlFilename(elementUrl) && (element.Content && !element.Content.name)) return undefined
			return isBlob(element.Content)
				? createWithRESTFromBlob.call(this, element, opts)
				: createWithRESTFromString.call(this, element, opts)
		})

		listReport('create', {
			...opts,
			name: this.name,
			box: getInstance(Box)(res),
			listUrl: this.listUrl,
			contextUrl: this.contextUrl
		})
		return res
	}

	async update(opts = {}) {
		const { contextUrl, listUrl } = this
		const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
		await this.cacheColumns()
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const { Content, Columns } = element
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			let spObject
			if (isUndefined(Content)) {
				spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, Columns))(
					this.getSPObject(elementUrl, clientContext).get_listItemAllFields()
				)
			} else {
				const fieldsToUpdate = {}
				const keys = Reflect.ownKeys(Columns)
				for (let i = 0; i < keys.length; i += 1) {
					const fieldName = keys[i]
					const field = Columns[fieldName]
					fieldsToUpdate[fieldName] = ifThen(isArray)([join(';#;#')])(field)
				}
				const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation)
				setFields({
					set_content: convertFileContent(Content),
					set_fieldValues: fieldsToUpdate
				})(binaryInfo)
				spObject = this.getSPObject(elementUrl, clientContext)
				spObject.saveBinary(binaryInfo)
				spObject = spObject.get_listItemAllFields()
			}
			return load(clientContext, spObject.get_file(), options)
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', options)
		return prepareResponseJSOM(result, options)
	}

	async delete(opts = {}) {
		const { contextUrl, listUrl } = this
		const { noRecycle } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const spObject = this.getSPObject(elementUrl, clientContext)
			methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	async copy(opts) {
		return copyOrMove.call(this, false, opts)
	}

	async move(opts) {
		return copyOrMove.call(this, true, opts)
	}

	getSPObject(elementUrl, clientContext) {
		const filename = getFilenameFromUrl(elementUrl)
		const folder = getFolderFromUrl(elementUrl)
		const rootSPFolder = this.getRootFolder(clientContext)
		return folder
			? getSPFolderByUrl(folder)(rootSPFolder)
				.get_files()
				.getByUrl(filename)
			: rootSPFolder
				.get_files()
				.getByUrl(filename)
	}

	getSPObjectCollection(elementUrl, clientContext) {
		const folder = getFolderFromUrl(popSlash(elementUrl))
		const rootSPFolder = this.getRootFolder(clientContext)
		return folder
			? getSPFolderByUrl(folder)(rootSPFolder).get_files()
			: rootSPFolder.get_files()
	}

	getRootFolder(clientContext) {
		return this.parent.getSPObject(this.listUrl, clientContext).get_rootFolder()
	}

	getRESTObject(elementUrl, listUrl, contextUrl) {
		let url = elementUrl
		if (/'/.test(elementUrl)) {
			url = url.replace(/'/, '%27%27')
		}
		return mergeSlashes(
			`${this.getRESTObjectCollection(elementUrl, listUrl, contextUrl)
			}/getbyurl('${getFilenameFromUrl(url)
			}')`
		)
	}

	getRESTObjectCollection(elementUrl, listUrl, contextUrl) {
		const folder = getFolderFromUrl(elementUrl)
		let folderUrl = ''
		if (folder) {
			const folders = mergeSlashes(folder).split('/')
			folderUrl = folders.reduce((acc, el) => `${acc}/folders/getbyurl('${el}')`, '')
		}

		return mergeSlashes(
			`/${contextUrl}/_api/web/lists/getbytitle('${listUrl}')/rootfolder${folderUrl}/files`
		)
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

	of(files) {
		return getInstance(this.constructor)(this.parent, files)
	}
}

export default getInstance(FileList)


// async function createWithJSOM(opts = {}) {
// 	let needToRetry
// 	let isError
// 	const { contextUrl, listUrl } = this
// 	const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }

// 	const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
// 		const { Content = '', Columns = {}, Overwrite = true } = element
// 		const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
// 		if (!hasUrlFilename(elementUrl)) return undefined
// 		const spObjects = this.getSPObjectCollection(elementUrl, clientContext)
// 		const fileCreationInfo = getInstanceEmpty(SP.FileCreationInformation)
// 		setFields({
// 			set_url: `/${contextUrl}/${listUrl}/${elementUrl}`,
// 			set_content: '',
// 			set_overwrite: Overwrite
// 		})(fileCreationInfo)
// 		const spObject = spObjects.add(fileCreationInfo)
// 		const fieldsToCreate = {}
// 		if (isObjectFilled(Columns)) {
// 			const props = Reflect.ownKeys(Columns)
// 			for (let i = 0; i < props.length; i += 1) {
// 				const prop = props[i]
// 				const fieldName = Columns[prop]
// 				const field = Columns[fieldName]
// 				fieldsToCreate[fieldName] = ifThen(isArray)([join(';#;#')])(field)
// 			}
// 		}
// 		const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation)
// 		setFields({
// 			set_content: convertFileContent(Content),
// 			set_fieldValues: fieldsToCreate
// 		})(binaryInfo)
// 		spObject.saveBinary(binaryInfo)
// 		return load(clientContext, spObject, options)
// 	})
// 	if (this.count) {
// 		for (let i = 0; i < clientContexts.length; i += 1) {
// 			const clientContext = clientContexts[i]
// 			await executorJSOM(clientContext).catch(async () => {
// 				isError = true
// 				needToRetry = await createUnexistedFolder.call(this)
// 			})
// 			if (needToRetry) break
// 		}
// 	}
// 	if (needToRetry) {
// 		return createWithJSOM.call(this, options)
// 	}
// 	if (isError) {
// 		throw new Error('can\'t create file(s)')
// 	} else {
// 		this.report('create', options)
// 		return prepareResponseJSOM(result, options)
// 	}
// }
