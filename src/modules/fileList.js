import axios from 'axios'
import {
	ACTION_TYPES,
	CACHE_RETRIES_LIMIT,
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
	typeOf,
	setItem,
	ifThen,
	join,
	getListRelativeUrl,
	listReport,
	deep3Iterator,
	deep2Iterator,
	switchCase,
	shiftSlash,
	isArrayFilled,
	pipe,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	constant,
	deep3IteratorREST,
	deep2IteratorREST,
	stringTest,
	isUndefined,
	isBlob,
	hasUrlFilename,
	removeEmptyFilenames,
	isObjectFilled,
	prependSlash,
	isNumberFilled
} from '../lib/utility'
import * as cache from '../lib/cache'

import web from './web'

// Internal

const NAME = 'file'

const getRequestDigest = contextUrl => axios({
	url: `${prependSlash(contextUrl)}/_api/contextinfo`,
	headers: {
		Accept: 'application/json; odata=verbose'
	},
	method: 'POST'
}).then(res => res.data.d.GetContextWebInformation.FormDigestValue)

const getSPObject = elementUrl => spObject => {
	const filename = getFilenameFromUrl(elementUrl)
	const folder = getFolderFromUrl(elementUrl)
	return folder
		? getSPFolderByUrl(folder)(spObject.get_rootFolder())
			.get_files()
			.getByUrl(filename)
		: spObject
			.get_rootFolder()
			.get_files()
			.getByUrl(filename)
}

const getSPObjectCollection = elementUrl => spObject => {
	const folder = getFolderFromUrl(popSlash(elementUrl))
	return folder
		? getSPFolderByUrl(folder)(spObject.get_rootFolder()).get_files()
		: spObject.get_rootFolder().get_files()
}

const getRESTObject = elementUrl => listUrl => contextUrl => mergeSlashes(
	`${getRESTObjectCollection(elementUrl)(listUrl)(contextUrl)}/getbyurl('${getFilenameFromUrl(elementUrl)}')`
)

const getRESTObjectCollection = elementUrl => listUrl => contextUrl => {
	const folder = getFolderFromUrl(elementUrl)
	const folderUrl = folder ? `/folders/getbyurl('${folder}')` : ''
	return mergeSlashes(`/${contextUrl}/_api/web/lists/getbytitle('${listUrl}')/rootfolder${folderUrl}/files`)
}

const liftFolderType = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context)
		const name = context.Content ? context.Content.name : undefined
		if (!context.Url) newContext.Url = context.ServerRelativeUrl || context.FileRef || name
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

	getCount() {
		return this.isArray ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0
	}
}

const iterator = instance => deep3Iterator({
	contextBox: instance.parent.parent.box,
	parentBox: instance.parent.box,
	elementBox: instance.box
	// bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
})

const iteratorREST = instance => deep3IteratorREST({
	contextBox: instance.parent.parent.box,
	parentBox: instance.parent.box,
	elementBox: instance.box
})

const iteratorParentREST = instance => deep2IteratorREST({
	contextBox: instance.parent.parent.box,
	elementBox: instance.parent.box
})

const report = instance => actionType => (opts = {}) => listReport({
	...opts,
	NAME,
	actionType,
	box: instance.box,
	listBox: instance.parent.box,
	contextBox: instance.parent.parent.box
})

const createNonexistedFolder = async instance => {
	const foldersToCreate = {}
	await iteratorREST(instance)(({ element }) => {
		foldersToCreate[element.Folder || getFolderFromUrl(element.Url)] = true
	})

	return iteratorParentREST(instance)(async ({ contextElement, element }) => web(contextElement.Url)
		.list(element.Url)
		.folder(Object.keys(foldersToCreate))
		.create({ silentInfo: true, expanded: true, view: ['Name'] })
		.then(() => {
			const cacheUrl = ['fileCreationRetries', instance.parent.parent.id]
			const retries = cache.get(cacheUrl)
			if (retries) {
				cache.set(retries - 1)(cacheUrl)
				return true
			}
			return false
		})
		.catch(err => {
			if (/already exists/.test(err.get_message())) return true
			return false
		}))
}

const createWithJSOM = instance => async (opts = {}) => {
	let needToRetry; let
		isError
	const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }

	const { clientContexts, result } = await iterator(instance)(
		({
			contextElement, clientContext, parentElement, element
		}) => {
			const { Content = '', Columns = {}, Overwrite = true } = element
			const listUrl = parentElement.Url
			const contextUrl = contextElement.Url
			const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
			if (!hasUrlFilename(elementUrl)) return undefined
			const contextSPObject = instance.parent.parent.getSPObject(clientContext)
			const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
			const spObjects = getSPObjectCollection(elementUrl)(listSPObject)
			const fileCreationInfo = getInstanceEmpty(SP.FileCreationInformation)
			setFields({
				set_url: `/${contextUrl}/${listUrl}/${elementUrl}`,
				set_content: '',
				set_overwrite: Overwrite
			})(fileCreationInfo)
			const spObject = spObjects.add(fileCreationInfo)
			const fieldsToCreate = {}
			if (isObjectFilled(Columns)) {
				const props = Reflect.ownKeys(Columns)
				for (let i = 0; i < props.length; i += 1) {
					const prop = props[i]
					const fieldName = Columns[prop]
					const field = Columns[fieldName]
					fieldsToCreate[fieldName] = ifThen(isArray)([join(';#;#')])(field)
				}
			}
			const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation)
			setFields({
				set_content: convertFileContent(Content),
				set_fieldValues: fieldsToCreate
			})(binaryInfo)
			spObject.saveBinary(binaryInfo)
			return load(clientContext)(spObject)(options)
		}
	)
	if (instance.box.getCount()) {
		await instance.parent.parent.box.chain(async el => {
			const clientContextByUrl = clientContexts[el.URL]
			for (let i = 0; i < clientContextByUrl.length; i += 1) {
				const clientContext = clientContextByUrl[i]
				await executorJSOM(clientContext)({ ...options, silentErrors: true }).catch(async () => {
					isError = true
					needToRetry = await createNonexistedFolder(instance)
				})
				if (needToRetry) break
			}
		})
	}
	if (needToRetry) {
		return createWithJSOM(instance)(options)
	}
	if (isError) {
		throw new Error('can\'t create file(s)')
	} else {
		report(instance)('create')(options)
		return prepareResponseJSOM(options)(result)
	}
}

const createWithRESTFromString = ({
	instance, contextUrl, listUrl, element
}) => async (opts = {}) => {
	let needToRetry; let
		isError
	const { needResponse } = opts
	const { Content = '', Overwrite = true, Columns } = element
	const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
	const filename = getFilenameFromUrl(elementUrl)
	const filesUrl = getRESTObjectCollection(elementUrl)(listUrl)(contextUrl)
	await axios({
		url: `${filesUrl}/add(url='${filename}',overwrite=${Overwrite})`,
		headers: {
			accept: 'application/json;odata=verbose',
			'content-type': 'application/json;odata=verbose',
			'X-RequestDigest': await getRequestDigest()
		},
		method: 'POST',
		data: Content
	}).catch(async err => {
		isError = true
		if (err.response.statusText === 'Not Found') {
			needToRetry = await createNonexistedFolder(instance)
		}
	})
	if (needToRetry) {
		return createWithRESTFromString({
			instance, contextUrl, listUrl, element
		})(opts)
	}
	if (isError) {
		throw new Error(`can't create file "${element.Url}" at ${contextUrl}/${listUrl}`)
	} else {
		let response
		if (Columns) {
			response = await web(contextUrl)
				.library(listUrl)
				.file({ Url: elementUrl, Columns })
				.update({ ...opts, silentInfo: true })
		} else if (needResponse) {
			response = await web(contextUrl)
				.library(listUrl)
				.file(elementUrl)
				.get(opts)
		}
		return response
	}
}

const createWithRESTFromBlob = ({
	instance, contextUrl, listUrl, element
}) => async (opts = {}) => {
	let isError
	let needToRetry
	const inputs = []
	const { needResponse, silent, silentErrors } = opts
	const {
		Content = '', Overwrite, OnProgress = identity, Folder = '', Columns
	} = element
	const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
	const folder = Folder || getFolderFromUrl(elementUrl)
	const filename = elementUrl ? getFilenameFromUrl(elementUrl) : Content.name
	const requiredInputs = {
		__REQUESTDIGEST: true,
		__VIEWSTATE: true,
		__EVENTTARGET: true,
		__EVENTVALIDATION: true,
		ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation: true,
		ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle: true
	}

	const listGUID = cache.get(['listGUIDs', contextUrl, listUrl])
	const listFormMatches = cache.get(['listFormMatches', contextUrl, listUrl])
	const inputRE = /<input[^<]*\/>/g
	let founds = inputRE.exec(listFormMatches)
	while (founds) {
		const item = founds[0]
		const id = item.match(/id="([^"]+)"/)[1]
		if (requiredInputs[id]) {
			switch (id) {
				case '__EVENTTARGET':
					inputs.push(item.replace(/value="[^"]*"/, 'value="ctl00$PlaceHolderMain$ctl03$RptControls$btnOK"'))
					break
				case 'ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation':
					inputs.push(item.replace(/value="[^"]*"/, `value="/${folder.replace(/^\//, '')}"`))
					break
				case 'ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle':
					inputs.push(Overwrite
						? item
						: item.replace(/checked="[^"]*"/, ''))
					break
				default:
					inputs.push(item)
					break
			}
		}
		founds = inputRE.exec(listFormMatches)
	}
	const form = window.document.createElement('form')
	form.innerHTML = join('')(inputs)
	const formData = new FormData(form)
	formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', Content, filename)

	const response = await axios({
		url: `/${contextUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
		method: 'POST',
		data: formData,
		onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
	})

	const errorMsgMatches = response.data.match(/id="ctl00_PlaceHolderMain_LabelMessage">([^<]*)<\/span>/)
	if (isArray(errorMsgMatches) && !silent && !silentErrors) console.error(errorMsgMatches[1])
	if (stringTest(/The selected location does not exist in this document library\./i)(response.data)) {
		isError = true
		needToRetry = await createNonexistedFolder(instance)
	}
	if (needToRetry) {
		return createWithRESTFromBlob({
			instance, contextUrl, listUrl, element
		})(opts)
	}
	if (isError) {
		throw new Error(`can't create file "${elementUrl}" at ${contextUrl}/${listUrl}`)
	} else {
		let res
		const list = web(contextUrl).library(listUrl)
		if (isObjectFilled(Columns)) {
			res = await list.file({ Url: elementUrl, Columns }).update({ ...opts, silentInfo: true })
		} else if (needResponse) {
			res = await list.file({ Url: elementUrl }).get(opts)
		}
		return res
	}
}

const copyOrMove = isMove => instance => async (opts = {}) => {
	await iteratorREST(instance)(async ({ contextElement, parentElement, element }) => {
		const { To, OnlyContent } = element
		const contextUrl = contextElement.Url
		const listUrl = parentElement.Url
		const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
		if (!hasUrlFilename(elementUrl)) return
		let targetWebUrl; let targetListUrl; let
			targetFileUrl
		if (isObject(To)) {
			targetWebUrl = To.WebUrl
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
		const spxSourceList = web(contextUrl).library(listUrl)
		const spxSourceFile = spxSourceList.file(elementUrl)
		const spxTargetList = web(targetWebUrl).library(targetListUrl)
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
			const listSPObject = instance.parent.getSPObject(listUrl)(instance.parent.parent.getSPObject(clientContext))
			const spObject = getSPObject(elementUrl)(listSPObject)
			const folder = getFolderFromUrl(targetFileUrl)
			if (folder) {
				await web(contextUrl)
					.library(listUrl)
					.folder(folder)
					.create({ silentInfo: true, expanded: true, view: ['Name'] })
					.catch(identity)
			}
			spObject[isMove ? 'moveTo' : 'copyTo'](mergeSlashes(`${targetListUrl}/${fullTargetFileUrl}`))
			await executeJSOM(clientContext)(spObject)(opts)
			await spxTargetList
				.file({ Url: targetFileUrl, Columns: existedColumnsToUpdate })
				.update({ silentInfo: true })
		} else {
			await spxTargetList
				.file({
					Url: fullTargetFileUrl,
					Content: await spxSourceList.file(elementUrl).get({ asBlob: true }),
					OnProgress: element.OnProgress,
					Overwrite: element.Overwrite,
					Columns: existedColumnsToUpdate
				})
				.create({ silentInfo: true })
			if (isMove) await spxSourceFile.delete()
		}
	})

	console.log(
		`${ACTION_TYPES[isMove ? 'move' : 'copy']} ${instance.parent.parent.box.getCount()
		* instance.parent.box.getCount()
		* instance.box.getCount()} ${NAME}(s)`
	)
}

const cacheListGUIDs = contextBox => elementBox => deep2Iterator({
	contextBox, elementBox
})(async ({ contextElement, element }) => {
	const contextUrl = contextElement.Url
	const listUrl = element.Url
	if (!cache.get(['listGUIDs', contextUrl, listUrl])) {
		const listProps = await web(contextUrl)
			.list(listUrl)
			.get({ view: 'Id' })
		cache.set(listProps.Id.toString())(['listGUIDs', contextUrl, listUrl])
	}
})

const cacheListFormMatches = contextBox => elementBox => deep2Iterator({
	contextBox, elementBox
})(async ({ contextElement, element }) => {
	const contextUrl = contextElement.Url
	const listUrl = element.Url
	if (!cache.get(['listFormMatches', contextUrl, listUrl])) {
		const listForms = await axios.get(
			`/${contextUrl}/_layouts/15/Upload.aspx?List={${cache.get(['listGUIDs', contextUrl, listUrl])}}`
		)
		cache.set(listForms.data.match(/<form(\w|\W)*<\/form>/))(['listFormMatches', contextUrl, listUrl])
	}
})

const cacheColumns = contextBox => elementBox => deep2Iterator({
	contextBox, elementBox
})(async ({ contextElement, element }) => {
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
})

// Inteface

export default parent => elements => {
	const instance = {
		box: getInstance(Box)(elements),
		parent
	}
	const instancedReport = report(instance)
	return {
		get: async (opts = {}) => {
			if (opts.asBlob) {
				const result = await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
					const contextUrl = contextElement.Url
					const listUrl = parentElement.Url
					const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
					return executorREST(contextUrl)({
						url: `${getRESTObject(elementUrl)(listUrl)(contextUrl)}/$value`,
						binaryStringResponseBody: true
					})
				})
				return prepareResponseREST(opts)(result)
			}
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			const { clientContexts, result } = await iterator(instance)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const listUrl = parentElement.Url
					const elementUrl = getListRelativeUrl(contextElement.Url)(listUrl)(element)
					const listSPObject = parent.getSPObject(listUrl)(parent.parent.getSPObject(clientContext))
					const spObject = isExists(elementUrl) && hasUrlTailSlash(elementUrl)
						? getSPObjectCollection(elementUrl)(listSPObject)
						: getSPObject(elementUrl)(listSPObject)
					return load(clientContext)(spObject)(options)
				}
			)
			await instance.parent.parent.box.chain(
				el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
			)
			return prepareResponseJSOM(options)(result)
		},

		create: async (opts = {}) => {
			await cacheListGUIDs(instance.parent.parent.box)(instance.parent.box)
			await cacheListFormMatches(instance.parent.parent.box)(instance.parent.box)
			const cacheUrl = ['fileCreationRetries', instance.parent.parent.id]
			if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
			const res = await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
				const contextUrl = contextElement.Url
				const listUrl = parentElement.Url
				const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
				if (!hasUrlFilename(elementUrl) && (element.Content && !element.Content.name)) return undefined
				return isBlob(element.Content)
					? createWithRESTFromBlob({
						instance, contextUrl, listUrl, element
					})(opts)
					: createWithRESTFromString({
						instance, contextUrl, listUrl, element
					})(opts)
			})
			instancedReport('create')(opts)
			return res
		},

		update: async (opts = {}) => {
			const options = opts.asItem ? { ...opts, view: ['ListItemAllFields'] } : { ...opts }
			await cacheColumns(instance.parent.parent.box)(instance.parent.box)
			const { clientContexts, result } = await iterator(instance)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const { Content, Columns } = element
					const listUrl = parentElement.Url
					const contextUrl = contextElement.Url
					const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
					if (!hasUrlFilename(elementUrl)) return undefined
					const contextSPObject = instance.parent.parent.getSPObject(clientContext)
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject)
					let spObject
					if (isUndefined(Content)) {
						spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, Columns))(
							getSPObject(elementUrl)(listSPObject).get_listItemAllFields()
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
						spObject = getSPObject(elementUrl)(listSPObject)
						spObject.saveBinary(binaryInfo)
						spObject = spObject.get_listItemAllFields()
					}
					return load(clientContext)(spObject.get_file())(options)
				}
			)
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(
					async el => Promise.all(clientContexts[el.Url].map(
						clientContext => executorJSOM(clientContext)(options)
					))
				)
			}
			instancedReport('update')(options)
			return prepareResponseJSOM(options)(result)
		},

		delete: async (opts = {}) => {
			const { noRecycle } = opts
			const { clientContexts, result } = await iterator(instance)(
				({
					contextElement, clientContext, parentElement, element
				}) => {
					const contextUrl = contextElement.Url
					const listUrl = parentElement.Url
					const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element)
					if (!hasUrlFilename(elementUrl)) return undefined
					const listSPObject = instance.parent.getSPObject(listUrl)(
						instance.parent.parent.getSPObject(clientContext)
					)
					const spObject = getSPObject(elementUrl)(listSPObject)
					methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
					return elementUrl
				}
			)
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(
					el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
				)
			}
			instancedReport(noRecycle ? 'delete' : 'recycle')(opts)
			return prepareResponseJSOM(opts)(result)
		},

		copy: copyOrMove(false)(instance),

		move: copyOrMove(true)(instance)
	}
}
