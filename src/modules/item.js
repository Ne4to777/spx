import { MD5 } from 'crypto-js'
import {
	MAX_ITEMS_LIMIT,
	CACHE_RETRIES_LIMIT,
	AbstractBox,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstanceEmpty,
	setItem,
	prop,
	methodEmpty,
	identity,
	listReport,
	getListRelativeUrl,
	isNull,
	isFilled,
	getInstance,
	switchCase,
	typeOf,
	shiftSlash,
	ifThen,
	isArrayFilled,
	map,
	constant,
	hasUrlTailSlash,
	isNumberFilled,
	isObjectFilled,
	getArray,
	popSlash,
	isDefined,
	arrayInit,
	deep1Iterator
} from '../lib/utility'

import * as cache from '../lib/cache'
import {
	getCamlView,
	craftQuery,
	camlLog,
	concatQueries
} from '../lib/query-parser'
import attachment from './attachment'

const getPagingColumnsStr = columns => {
	if (!columns) return ''
	let str = ''
	const keys = Reflect.keys(columns)
	for (let i = 0; i < keys.length; i += 1) {
		let valueStr = value
		const name = keys[i]
		const value = columns[name]
		switch (typeOf(value)) {
			case 'object':
				if (value.get_lookupValue) valueStr = value.get_lookupValue()
				break
			case 'date':
				valueStr = value.toISOString()
				break
			case 'boolean':
				valueStr = value ? 1 : 0
				break
			default:
			// default
		}
		if (value !== 'ID' && value !== undefined) str += `&p_${name}=${encodeURIComponent(valueStr)}`
	}
	return str
}

const getTypedSPObject = (typeStr, listUrl) => (element, parentElement) => {
	let camlQuery = new SP.CamlQuery()
	const { ID } = element
	const contextUrl = parentElement.get_context().get_url()
	switch (typeOf(ID)) {
		case 'number': {
			const spObject = parentElement.getItemById(ID)
			return spObject
		}
		case 'string':
			if (hasUrlTailSlash(ID)) {
				camlQuery.set_folderServerRelativeUrl(`${contextUrl}/${popSlash(ID)}`)
			} else {
				camlQuery.set_viewXml(getCamlView(ID))
			}
			break
		default:
			if (element.set_viewXml) {
				camlQuery = element
			} else if (element.get_lookupId) {
				const spObject = parentElement.getItemById(element.get_lookupId())
				return spObject
			} else {
				const { Folder, Page } = element
				if (Folder) {
					let fullUrl = ''
					const folderUrl = `${typeStr}${listUrl}/${Folder}`
					if (contextUrl === '/') {
						fullUrl = folderUrl
					} else {
						const isFullUrl = new RegExp(parentElement.get_context().get_url(), 'i').test(Folder)
						fullUrl = isFullUrl ? Folder : folderUrl
					}
					camlQuery.set_folderServerRelativeUrl(fullUrl)
				}
				camlQuery.set_viewXml(getCamlView(element))
				if (Page) {
					const { IsPrevious, Id, Columns } = Page
					if (Id) {
						const position = getInstanceEmpty(SP.ListItemCollectionPosition)
						position.set_pagingInfo(
							`${IsPrevious ? 'PagedPrev=TRUE&' : ''}Paged=TRUE&p_ID=${Id}${getPagingColumnsStr(Columns)}`
						)
						camlQuery.set_listItemCollectionPosition(position)
					}
				}
			}
	}
	const spObject = parentElement.getItems(camlQuery)
	spObject.camlQuery = camlQuery
	return spObject
}

const liftItemType = switchCase(typeOf)({
	object: item => {
		const newItem = Object.assign({}, item)
		if (!item.Url) newItem.Url = item.ID
		if (!item.ID) newItem.ID = item.Url
		return newItem
	},
	string: (item = '') => {
		const url = shiftSlash(item)
		return {
			Url: url,
			ID: url
		}
	},
	number: item => ({
		ID: item,
		Url: item
	}),
	default: () => ({
		ID: undefined,
		Url: undefined
	})
})

const hasObjectProperties = o => {
	let newObject
	if (o.Columns) {
		newObject = Object.assign({}, o.Columns)
	} else {
		newObject = Object.assign({}, o)
		delete newObject.Folder
		delete newObject.ID
		delete newObject.Url
	}
	return !!Object.keys(newObject).length
}

class Box extends AbstractBox {
	constructor(value) {
		super(value)
		this.joinProp = 'ID'
		this.value = this.isArray
			? ifThen(isArrayFilled)([map(liftItemType), constant([liftItemType()])])(value)
			: liftItemType(value)
	}

	getCount(actionType) {
		if (this.isArray) {
			return actionType === 'create' ? this.value.filter(hasObjectProperties).length : this.value.length
		}
		return actionType === 'create' ? (hasObjectProperties(this.value) ? 1 : 0) : 1
	}
}

class Item {
	constructor(parent, items) {
		this.name = 'item'
		this.parent = parent
		this.box = getInstance(Box)(items)
		this.listUrl = parent.box.head().Url
		this.getSPObject = getTypedSPObject(parent.name === 'list' ? 'Lists/' : '', this.listUrl)
		this.getListSPObject = this.parent.getSPObject
		this.getContextSPObject = this.parent.parent.getSPObject
		this.iterator = deep1Iterator({
			contextUrl: parent.contextUrl,
			elementBox: this.box
		})
	}

	async	get(opts = {}) {
		const { showCaml } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listUrl, contextSPObject)
			const spObject = this.getSPObject(element, listSPObject)
			if (showCaml && spObject.camlQuery) camlLog(spObject.camlQuery.get_viewXml())
			return load(clientContext)(spObject)(opts)
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		return prepareResponseJSOM(opts)(result)
	}

	async create(opts = {}) {
		if (this.parent.name !== 'list') {
			throw new Error('There is no "create" method in "item" module. Use "file" module instead of "item".')
		}
		const { contextUrl, listUrl } = this
		await this.cacheColumns()
		const cacheUrl = ['itemCreationRetries', this.parent.parent.id]
		if (!isNumberFilled(cache.get(cacheUrl))) cache.set(CACHE_RETRIES_LIMIT)(cacheUrl)
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const { Folder, Columns } = element
			let folder = Folder
			let newElement
			if (isObjectFilled(Columns)) {
				if (Columns.Folder) folder = Columns.Folder
				newElement = { ...Columns }
			} else {
				newElement = { ...element }
				delete newElement.ID
				delete newElement.Url
				delete newElement.Folder
			}
			if (!isObjectFilled(newElement)) return undefined
			if (folder) {
				const folderUrl = getListRelativeUrl(contextUrl)(listUrl)({ Folder: folder })
				itemCreationInfo.set_folderUrl(`/${contextUrl}/Lists/${listUrl}/${folderUrl}`)
			}
			const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(newElement)(
				listSPObject.addItem(itemCreationInfo)
			)
			return load(clientContext)(spObject)(opts)
		})
		let needToRetry; let
			isError
		if (this.box.getCount()) {
			for (let i = 0; i < clientContexts.length; i += 1) {
				const clientContext = clientContexts[i]
				await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
					if (/There is no file with URL/.test(err.get_message())) {
						const foldersToCreate = {}
						await this.iterator(({ element }) => {
							const { Folder, Columns = {} } = element
							const elementUrl = getListRelativeUrl(this.contextUrl)(this.listUrl)({
								Folder: Folder || Columns.Folder
							})
							foldersToCreate[elementUrl] = true
						})

						const res = await this
							.parent
							.folder(Object.keys(foldersToCreate))
							.create({ expanded: true, view: ['Name'] })
							.then(() => {
								const retries = cache.get(cacheUrl)
								if (retries) {
									cache.set(retries - 1)(cacheUrl)
									return true
								}
								return false
							})
							.catch(error => {
								console.log(error)
								if (/already exists/.test(error.get_message())) return true
								return false
							})
						if (res) needToRetry = true
					} else {
						throw err
					}
					isError = true
				})
				if (needToRetry) break
			}
		}
		if (needToRetry) {
			return this.create(opts)
		}
		if (!isError) {
			this.report('create', opts)
			return prepareResponseJSOM(opts)(result)
		}
		return undefined
	}

	async	update(opts) {
		const { contextUrl, listUrl } = this
		await this.cacheColumns()
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			if (!element.Url) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const elementNew = Object.assign({}, element)
			delete elementNew.ID
			const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(elementNew)(
				this.getSPObject(element, listSPObject)
			)
			return load(clientContext)(spObject)(opts)
		})
		if (isFilled(result)) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('update')(opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	updateByQuery(opts) {
		const { result } = await this.iterator(async ({ element }) => {
			const { Folder, Query, Columns = {} } = element
			if (Query === undefined) throw new Error('Query is missed')
			const items = await this.of({ Folder, Query }).get({ ...opts, expanded: false })
			if (items.length) {
				const itemsToUpdate = items.map(item => {
					const columns = { ...Columns }
					columns.ID = item.ID
					return columns
				})
				return this.of(itemsToUpdate).update(opts)
			}
			return undefined
		})
		return result
	}

	async	delete(opts = {}) {
		const { noRecycle, isSerial } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = element.Url
			if (!elementUrl) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listUrl, contextSPObject)
			const spObject = this.getSPObject(element, listSPObject)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementUrl
		})

		if (this.box.getCount()) {
			if (isSerial) {
				for (let i = 0; i < clientContexts.length; i += 1) {
					await executorJSOM(clientContexts[i])(opts)
				}
			} else {
				await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
			}
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	deleteByQuery(opts) {
		const { result } = await this.iterator(async ({ element }) => {
			if (element === undefined) throw new Error('Query is missed')
			const items = await this.of(element).get(opts)
			if (items.length) {
				await this.of(items.map(prop('ID'))).delete({ isSerial: true })
				return items
			}
			return undefined
		})
		return result
	}

	async	getDuplicates() {
		const items = await this
			.of({
				Scope: 'allItems',
				Limit: MAX_ITEMS_LIMIT
			})
			.get({
				view: ['ID', 'FSObjType', 'Modified', ...this.box.join()]
			})
		const itemsMap = new Map()
		const getHashedColumnName = itemData => {
			const values = []
			this.box.chain(column => {
				const value = itemData[column.ID]
				if (isDefined(value)) values.push(value.get_lookupId ? value.get_lookupId() : value)
			})
			return MD5(values.join('.')).toString()
		}
		for (let i = 0; i < items.length; i += 1) {
			const item = items[i]
			const hashedColumnName = getHashedColumnName(item)
			if (hashedColumnName !== undefined) {
				if (!itemsMap.has(hashedColumnName)) itemsMap.set(hashedColumnName, [])
				itemsMap.get(hashedColumnName).push(item)
			}
		}

		let duplicatedsSorted = []
		for (let i = 0; i < itemsMap.length; i += 1) {
			const [, duplicateds] = itemsMap[i]
			if (duplicateds.length > 1) {
				duplicatedsSorted = duplicatedsSorted.concat([
					duplicateds.sort((a, b) => a.Modified.getTime() - b.Modified.getTime())
				])
			}
		}
		return duplicatedsSorted.map(arrayInit)
	}

	async	deleteDuplicates(opts = {}) {
		const duplicatedsFiltered = await this.getDuplicates()
		await this
			.of([...duplicatedsFiltered.reduce((acc, el) => acc.concat(el.map(prop('ID'))), [])])
			.delete({ ...opts, isSerial: true })
	}

	async getEmpties(opts) {
		const columns = this.box.value.map(prop('ID'))
		const { result } = await this.iterator(async ({ element }) => this
			.of({
				Query: craftQuery({ operator: 'isnull', columns }),
				Scope: 'allItems',
				Limit: MAX_ITEMS_LIMIT,
				Folder: element.Folder
			})
			.get(opts))
		return result
	}

	async	deleteEmpties(opts) {
		const columns = this.box.value.map(prop('ID'))
		const { result } = await this
			.of(await this.of(columns).getEmpties(opts)).map(prop('ID'))
			.delete({ isSerial: true })
		return result
	}

	async merge(opts) {
		const { contextUrl, listUrl } = this
		this.iterator(async ({ element }) => {
			const {
				Key,
				Forced,
				To,
				From,
				Mediator = () => identity
			} = element
			if (!Key) throw new Error('Key is missed')
			const sourceColumn = From.Column
			if (!sourceColumn) throw new Error('From Column is missed')
			const targetColumn = To.Column
			if (!targetColumn) throw new Error('To Column is missed')
			const [sourceItems, targetItems] = await Promise.all([
				this.parent.parent
					.of(From.WebUrl || contextUrl)[this.parent.name](From.List || listUrl)
					.item({
						Query: concatQueries()([`${Key} IsNotNull`, From.Query]),
						Scope: 'allItems',
						Limit: MAX_ITEMS_LIMIT
					})
					.get({
						view: ['ID', Key, sourceColumn]
					}),
				this
					.of({
						Query: concatQueries()([`${Key} IsNotNull`, To.Query]),
						Scope: 'allItems',
						Limit: MAX_ITEMS_LIMIT
					})
					.get({
						view: ['ID', Key, targetColumn],
						groupBy: Key
					})
			])
			const itemsToMerge = []
			await Promise.all(
				sourceItems.map(async sourceItem => {
					const targetItemGroup = targetItems[sourceItem[Key]]
					if (targetItemGroup) {
						const targetItem = targetItemGroup[0]
						const itemToMerge = { ID: targetItem.ID }
						if (isNull(targetItem[targetColumn]) || Forced) {
							itemToMerge[targetColumn] = await Mediator(sourceColumn)(sourceItem[sourceColumn])
							itemsToMerge.push(itemToMerge)
						}
					}
				})
			)
			return itemsToMerge.length && this.of(itemsToMerge).update(opts)
		})
	}

	async	erase(opts) {
		const { result } = await this.iterator(({ element }) => {
			const { Query = '', Folder, Columns } = element
			if (!Columns) return undefined
			const columns = getArray(Columns)
			return this
				.of({
					Folder,
					Query,
					Columns: columns.reduce((acc, el) => {
						acc[el] = null
						return acc
					}, {})
				})
				.updateByQuery(opts)
		})
		return prepareResponseJSOM(opts)(result)
	}

	async	createDiscussion(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listUrl, contextSPObject)
			const skippedProps = {
				Url: true,
				ID: true
			}
			const spObject = SP.Utilities.Utility.createNewDiscussion(clientContext, listSPObject, element.Title)
			const keys = Reflect.ownKeys(element)
			for (let i = 0; i < keys.length; i += 1) {
				const columnName = keys[i]
				if (!skippedProps[columnName]) spObject.set_item(columnName, element[columnName])
			}
			spObject.update()
			return spObject
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		this.report('create', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	createReply(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const { Columns } = element
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listUrl, contextSPObject)
			const spItemObject = this.getSPObject(element, listSPObject)
			const spObject = SP.Utilities.Utility.createNewDiscussionReply(clientContext, spItemObject)
			const keys = Reflect.ownKeys(Columns)
			for (let i = 0; i < keys.length; i += 1) {
				const columnName = keys[i]
				spObject.set_item(columnName, Columns[columnName])
			}
			spObject.update()
			return spObject
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		this.report('create', opts)
		return prepareResponseJSOM(opts)(result)
	}

	attachment(attachments) {
		return attachment(this, attachments)
	}

	report(actionType, opts = {}) {
		listReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			listBox: this.parent.box,
			contextBox: this.parent.parent.box
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
					groupBy: 'InternalName'
				})
			cache.set(columns)(['columns', contextUrl, listUrl])
		}
	}

	of(items) {
		return getInstance(this.constructor)(this.parent, items)
	}
}

export default getInstance(Item)
