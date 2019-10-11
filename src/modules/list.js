/* eslint class-methods-use-this:0 */
import {
	FILE_LIST_TEMPLATES,
	AbstractBox,
	pipe,
	methodEmpty,
	methodI,
	ifThen,
	constant,
	prepareResponseJSOM,
	getClientContext,
	load,
	executorJSOM,
	getInstanceEmpty,
	setFields,
	overstep,
	hasUrlTailSlash,
	getTitleFromUrl,
	isExists,
	isString,
	webReport,
	switchCase,
	typeOf,
	getInstance,
	isGUID,
	method,
	removeEmptiesByProp,
	removeDuplicatedProp,
	getListRelativeUrl,
	deep1Iterator,
	stringTest,
	shiftSlash,
	popSlash,
	getPermissionMasks
} from '../lib/utility'
import column from './column'
import folder from './folderList'
import file from './fileList'
import item from './item'
import pager from './pager'
import { getCamlQuery, getCamlScope } from '../lib/query-parser'

const KEY_PROP = 'Title'

const arrayValidator = pipe([removeEmptiesByProp(KEY_PROP), removeDuplicatedProp(KEY_PROP)])

const lifter = switchCase(typeOf)({
	object: list => {
		const newList = Object.assign({}, list)
		const { Url, EntityTypeName, Title } = list
		newList[KEY_PROP] = EntityTypeName || Url || Title
		if (stringTest(/\//)(Url)) {
			if (Url !== '/') {
				newList[KEY_PROP] = popSlash(getTitleFromUrl(Url))
			}
			newList.Url = hasUrlTailSlash(Url) ? '/' : shiftSlash(Url) || '/'
		}
		return newList
	},
	string: (listUrl) => {
		const newList = {
			Url: undefined,
			Title: listUrl
		}
		if (stringTest(/\//)(listUrl)) {
			newList.Title = getTitleFromUrl(listUrl)
			newList.Url = hasUrlTailSlash(listUrl) ? '/' : shiftSlash(listUrl) || '/'
		}
		return newList
	},
	default: () => ({
		Url: '/',
		Title: undefined
	})
})

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter, arrayValidator)
		this.prop = KEY_PROP
		this.joinProp = KEY_PROP
	}
}

class List {
	constructor(parent, lists) {
		this.name = 'list'
		this.prop = KEY_PROP
		this.joinProp = KEY_PROP
		this.parent = parent
		this.contextUrl = parent.box.getHeadPropValue()
		this.box = getInstance(Box)(lists)
		this.count = parent.box.getCount()
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box
		})
	}

	async get(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const isCollection = hasUrlTailSlash(element.Url)
			const spObject = isCollection
				? this.getSPObjectCollection(clientContext)
				: this.getSPObject(element.Title, clientContext)
			return load(clientContext, spObject, opts)
		})

		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}

	async create(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!title) return undefined
			const spObject = pipe([
				getInstanceEmpty,
				setFields({
					set_title: title || undefined,
					set_templateType: element.BaseTemplate
						|| SP.ListTemplateType[element.TemplateType
						|| 'genericList'],
					set_url: element.Url === '/' ? title : element.Url,
					set_templateFeatureId: element.TemplateFeatureId,
					set_customSchemaXml: element.CustomSchemaXml,
					set_dataSourceProperties: element.DataSourceProperties,
					set_documentTemplateType: element.DocumentTemplateType,
					set_quickLaunchOption: element.QuickLaunchOption
				}),
				methodI('add')(this.getSPObjectCollection(clientContext)),
				overstep(
					setFields({
						set_contentTypesEnabled: element.ContentTypesEnabled,
						set_defaultContentApprovalWorkflowId: element.DefaultContentApprovalWorkflowId,
						set_defaultDisplayFormUrl: element.DefaultDisplayFormUrl,
						set_defaultEditFormUrl: element.DefaultEditFormUrl,
						set_defaultNewFormUrl: element.DefaultNewFormUrl,
						set_description: element.Description,
						set_direction: element.Direction,
						set_draftVersionVisibility: element.DraftVersionVisibility,
						set_enableAttachments: element.EnableAttachments || false,
						set_enableFolderCreation: element.EnableFolderCreation === undefined
							? true
							: element.EnableFolderCreation,
						set_enableMinorVersions: element.EnableMinorVersions,
						set_enableModeration: element.EnableModeration,
						set_enableVersioning: element.EnableVersioning === undefined
							? true
							: element.EnableVersioning,
						set_forceCheckout: element.ForceCheckout,
						set_hidden: element.Hidden,
						set_imageUrl: element.ImageUrl,
						set_irmEnabled: element.IrmEnabled,
						set_irmExpire: element.IrmExpire,
						set_irmReject: element.IrmReject,
						set_isApplicationList: element.IsApplicationList,
						set_lastItemModifiedDate: element.LastItemModifiedDate,
						set_majorVersionLimit: element.EnableVersioning ? element.MajorVersionLimit : undefined,
						set_multipleDataList: element.MultipleDataList,
						set_noCrawl: element.NoCrawl === undefined ? true : element.NoCrawl,
						set_objectVersion: element.ObjectVersion,
						set_onQuickLaunch: element.OnQuickLaunch,
						set_validationFormula: element.ValidationFormula,
						set_validationMessage: element.ValidationMessage
					})
				),
				overstep(
					ifThen(constant(!element.BaseTemplate && !FILE_LIST_TEMPLATES[element.TemplateType]))([
						setFields({
							set_documentTemplateUrl: element.DocumentTemplateUrl,
							MajorWithMinorVersionsLimit: element.EnableVersioning
								? element.MajorWithMinorVersionsLimit
								: undefined
						})
					])
				),
				overstep(methodEmpty('update'))
			])(SP.ListCreationInformation)
			return load(clientContext, spObject, opts)
		})
		if (this.count) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				await executorJSOM(clientContexts[i])
			}
		}
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	async update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!title) return undefined
			const spObject = pipe([
				setFields({
					set_title: title,
					set_enableFolderCreation: element.EnableFolderCreation,
					set_contentTypesEnabled: element.ContentTypesEnabled,
					set_defaultContentApprovalWorkflowId: element.DefaultContentApprovalWorkflowId,
					set_defaultDisplayFormUrl: element.DefaultDisplayFormUrl,
					set_defaultEditFormUrl: element.DefaultEditFormUrl,
					set_defaultNewFormUrl: element.DefaultNewFormUrl,
					set_description: element.Description,
					set_direction: element.Direction,
					set_draftVersionVisibility: element.DraftVersionVisibility,
					set_documentTemplateUrl: element.DocumentTemplateUrl,
					set_enableAttachments: element.EnableAttachments,
					set_enableMinorVersions: element.EnableMinorVersions,
					set_enableModeration: element.EnableModeration,
					set_enableVersioning: element.EnableVersioning,
					set_forceCheckout: element.ForceCheckout,
					set_hidden: element.Hidden,
					set_imageUrl: element.ImageUrl,
					set_irmEnabled: element.IrmEnabled,
					set_irmExpire: element.IrmExpire,
					set_irmReject: element.IrmReject,
					set_isApplicationList: element.IsApplicationList,
					set_lastItemModifiedDate: element.LastItemModifiedDate,
					set_majorVersionLimit: element.MajorVersionLimit,
					set_majorWithMinorVersionsLimit: element.MajorWithMinorVersionsLimit,
					set_multipleDataList: element.MultipleDataList,
					set_noCrawl: element.NoCrawl,
					set_objectVersion: element.ObjectVersion,
					set_onQuickLaunch: element.OnQuickLaunch,
					set_validationFormula: element.ValidationFormula,
					set_validationMessage: element.ValidationMessage
				}),
				overstep(methodEmpty('update'))
			])(this.getSPObject(title, clientContext))
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
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!title) return undefined
			const spObject = this.getSPObject(title, clientContext)
			methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return title
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	async cloneLayout() {
		console.log('cloning layout in progress...')
		await this.iterator(async ({ element }) => {
			let targetListUrl
			let targetWebUrl
			const { Title, Url, To } = element
			if (isString(To)) {
				targetWebUrl = this.contextUrl
				targetListUrl = To
			} else {
				targetWebUrl = To.WebUrl || this.contextUrl
				targetListUrl = To.ListUrl
			}
			const sourceListUrl = Url || Title
			if (!targetListUrl) throw new Error('Target listUrl is missed')
			if (!sourceListUrl) throw new Error('Source list Title is missed')
			const targetTitle = getTitleFromUrl(targetListUrl)
			const targetWeb = this.parent.of(targetWebUrl)
			const sourceWeb = this.parent.of(this.contextUrl)
			const targetList = targetWeb.list(targetListUrl)
			const sourceList = sourceWeb.list(sourceListUrl)
			await targetList.get({ silent: true }).catch(async () => {
				const newListData = Object.assign({}, await sourceList.get())
				newListData.Title = targetTitle
				newListData.Url = targetTitle
				newListData.EntityTypeName = targetTitle
				await targetWeb.list(newListData).create()
			})
			const [sourceColumns, targetColumns] = await Promise.all([
				sourceList.column().get(),
				targetList.column().get({ groupBy: 'InternalName' })
			])
			const columnsToCreate = sourceColumns.reduce((acc, sourceColumn) => {
				const targetColumn = targetColumns[sourceColumn.InternalName]
				if (!targetColumn && !sourceColumn.FromBaseType) acc.push(Object.assign({}, sourceColumn))
				return acc
			}, [])
			return columnsToCreate.length ? targetList.column(columnsToCreate).create() : undefined
		})
		console.log('cloning layout done!')
	}

	async clone(opts) {
		const columnsToExclude = {
			Attachments: true,
			MetaInfo: true,
			FileLeafRef: true,
			FileDirRef: true,
			FileRef: true,
			Order: true
		}
		console.log('cloning in progress...')
		await this.iterator(async ({ contextElement, element }) => {
			const contextUrl = contextElement.Url
			let targetWebUrl; let
				targetListUrl
			const { Title, To, Url } = element
			if (isString(To)) {
				targetWebUrl = contextUrl
				targetListUrl = To
			} else {
				targetWebUrl = To.WebUrl
				targetListUrl = To.ListUrl
			}
			const sourceListUrl = Title || Url
			if (!targetWebUrl) throw new Error('Target webUrl is missed')
			if (!targetListUrl) throw new Error('Target listUrl is missed')
			if (!sourceListUrl) throw new Error('Source list Title is missed')

			const targetTitle = To.Title || getTitleFromUrl(targetListUrl)

			const targetSPX = this.web(targetWebUrl)
			const sourceSPX = this.web(contextUrl)
			const targetSPXList = targetSPX.list(targetTitle)
			const sourceSPXList = sourceSPX.list(sourceListUrl)

			await sourceSPX.list(element).cloneLayout()
			console.log('cloning items...')
			const [sourceListData, sourceColumnsData, sourceItemsData] = await Promise.all([
				sourceSPXList.get(),
				sourceSPXList.column().get({ groupBy: 'InternalName', view: ['ReadOnlyField', 'InternalName'] }),
				sourceSPXList.item({ Query: '', Scope: 'allItems' }).get()
			])

			if (sourceListData.BaseType) {
				for (let i = 0; i < sourceItemsData.length; i += 1) {
					const fileItem = sourceItemsData[i]
					await sourceSPX
						.library(sourceListUrl)
						.file({
							Url: getListRelativeUrl(contextUrl)(sourceListUrl)({ Url: fileItem.FileRef }),
							To: {
								WebUrl: targetWebUrl,
								ListUrl: targetListUrl
							}
						})
						.copy(opts)
				}
			} else {
				const itemsToCreate = sourceItemsData.map(itemToCreate => {
					const newItem = {}
					const folderForItem = getListRelativeUrl(contextUrl)(sourceListUrl)(
						{ Url: itemToCreate.FileDirRef }
					)
					if (folderForItem) newItem.Folder = folderForItem
					const keys = Reflect.ownKeys(itemToCreate)
					for (let i = 0; i < keys.length; i += 1) {
						const prop = keys[i]
						const value = itemToCreate[prop]
						const sourceColumns = sourceColumnsData[prop]
						if (sourceColumns) {
							const sourceColumn = sourceColumns[0]
							if (
								!sourceColumn.ReadOnlyField && !columnsToExclude[prop] && isExists(value)
							) newItem[prop] = value
						}
					}
					return newItem
				})
				return targetSPXList.item(itemsToCreate).create()
			}
			return undefined
		})
		console.log('cloning is complete!')
	}

	async clear(opts) {
		console.log('clearing in progress...')
		await this
			.item({ Query: '' })
			.deleteByQuery(opts)
		console.log('clearing is complete!')
	}

	async getAggregations() {
		return this.iterator(async ({ contextElement, element }) => {
			const contextUrl = contextElement.Url
			let scopeStr = ''
			let fieldRefs = ''
			let caml = ''
			const aggregations = {}
			const {
				Title,
				Columns,
				Scope = 'all',
				Query
			} = element
			const clientContext = getClientContext(contextUrl)
			const list = this.getSPObject(Title, clientContext)
			const keys = Reflect.ownKeys(Columns)
			for (let i = 0; i < keys.length; i += 1) {
				const columnName = keys[i]
				fieldRefs += `<FieldRef Name="${columnName}" Type="${Columns[columnName]}"/>`
				aggregations[columnName] = 0
			}
			if (Scope) {
				scopeStr = getCamlScope()
			}
			if (Query) caml = `<Query><Where>${getCamlQuery(Query)}</Where></Query>`
			const aggregationsQuery = list.renderListData(
				`<View${scopeStr}>${caml}<Aggregations>${fieldRefs}</Aggregations></View>`
			)
			await executorJSOM(clientContext)
			const aggregationsData = JSON.parse(aggregationsQuery.get_value()).Row[0]
			const aggregationKeys = Reflect.ownKeys(aggregationsData)
			for (let i = 0; i < aggregationKeys.length; i += 1) {
				const name = aggregationKeys[i]
				const columnName = name.split('.')[0]
				if (Reflect.hasOwnProperty.call(Columns, columnName)) {
					aggregations[columnName] = Number(aggregationsData[name])
				}
			}
			return aggregations
		})
	}

	async getPermissions() {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const spObject = this.getSPObject(element[KEY_PROP], clientContext)
			return load(clientContext, spObject, { view: 'EffectiveBasePermissions' })
		})

		await Promise.all(clientContexts.map(executorJSOM))

		const allPermissionMasks = getPermissionMasks()
		return Object.keys(allPermissionMasks).reduce((acc, el) => {
			acc[el] = result.get_effectiveBasePermissions().has(allPermissionMasks[el])
			return acc
		}, {})
	}

	async breakRoleInheritance() {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!title) return undefined
			const spObject = this.getSPObject(title, clientContext)
			methodEmpty('breakRoleInheritance')(spObject)
			return title
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}

		return result
	}

	async resetRoleInheritance() {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!title) return undefined
			const spObject = this.getSPObject(title, clientContext)
			methodEmpty('resetRoleInheritance')(spObject)
			return title
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}

		return result
	}

	column(elements) {
		return column(this, elements)
	}

	folder(elements) {
		return folder(this, elements)
	}

	item(elements) {
		return item(this, elements)
	}

	pager(params) {
		return pager(this, params)
	}

	getSPObject(elementTitle, clientContext) {
		return this.getSPObjectCollection(clientContext)[
			isGUID(elementTitle)
				? 'getById'
				: 'getByTitle'
		](elementTitle)
	}

	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(clientContext).get_lists()
	}

	report(actionType, opts = {}) {
		webReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			contextUrl: this.contextUrl
		})
	}

	of(lists) {
		return getInstance(this.constructor)(this.parent, lists)
	}
}

class Library extends List {
	constructor(parent, libraries) {
		super(parent, libraries)
		this.name = 'library'
	}

	file(elements) {
		return file(this, elements)
	}
}

export default (parent, elements, name) => getInstance(name === 'library' ? Library : List)(parent, elements)
