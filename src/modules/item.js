/* eslint-disable no-restricted-syntax */
/* eslint no-shadow:0 */
import { MD5 } from 'crypto-js'
import {
	MAX_ITEMS_LIMIT,
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
	switchType,
	typeOf,
	hasUrlTailSlash,
	isObjectFilled,
	getArray,
	popSlash,
	isExists,
	arrayInit,
	deep1Iterator,
	isArray
} from '../lib/utility'

import * as cache from '../lib/cache'
import {
	getCamlView,
	craftQuery,
	camlLog,
	concatQueries
} from '../lib/query-parser'
import attachment from './attachment'

const KEY_PROP = 'ID'

const getPagingColumnsStr = columns => {
	if (!columns) return ''
	let str = ''
	const keys = Reflect.ownKeys(columns)
	for (let i = 0; i < keys.length; i += 1) {
		const name = keys[i]
		const value = columns[name]
		let valueStr = value
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


const lifter = switchType({
	object: item => Object.assign({}, item),
	string: (item = '') => ({
		[KEY_PROP]: item
	}),
	number: item => ({
		[KEY_PROP]: item
	}),
	default: () => ({
		[KEY_PROP]: undefined
	})
})

const hasObjectProperties = o => {
	let newObject
	if (o.Columns) {
		newObject = Object.assign({}, o.Columns)
	} else {
		newObject = Object.assign({}, o)
		delete newObject.Folder
		delete newObject[KEY_PROP]
	}
	return !!Object.keys(newObject).length
}

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter)
		this.prop = KEY_PROP
		this.joinProp = KEY_PROP
	}

	getCount(actionType) {
		if (isArray(this.value)) {
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
		this.contextUrl = parent.parent.box.getHeadPropValue()
		this.listUrl = parent.box.getHeadPropValue()
		this.listFolder = parent.name === 'list' ? 'Lists/' : ''
		this.getListSPObject = parent.getSPObject.bind(parent)
		this.iterator = deep1Iterator({
			contextUrl: parent.contextUrl,
			elementBox: this.box
		})
	}

	async get(opts = {}) {
		const { showCaml } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const spObject = this.getSPObject(element, clientContext)
			if (showCaml && spObject.camlQuery) camlLog(spObject.camlQuery.get_viewXml())
			return load(clientContext, spObject, opts)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, opts)
	}

	async create(opts = {}) {
		if (this.parent.name !== 'list') {
			throw new Error('There is no "create" method in "item" module. Use "file" module instead of "item".')
		}
		const { contextUrl, listUrl } = this
		await this.cacheColumns()

		const foldersToCreate = this.box.reduce(acc => el => {
			const { Folder } = el
			if (Folder) acc.push(Folder)
			return acc
		})

		if (foldersToCreate.length) {
			await this.parent.folder(foldersToCreate).get({ view: 'ServerRelativeUrl' }).catch(async () => {
				await this.parent.folder(foldersToCreate).create({ silent: true })
			})
		}

		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation)
			const listSPObject = this.getListSPObject(listUrl, clientContext)
			const { Folder, Columns } = element
			let folder = Folder
			let newElement
			if (isObjectFilled(Columns)) {
				if (Columns.Folder) folder = Columns.Folder
				newElement = { ...Columns }
			} else {
				newElement = { ...element }
				delete newElement[KEY_PROP]
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
			return load(clientContext, spObject, opts)
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
		const { contextUrl, listUrl } = this
		await this.cacheColumns()
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			if (!element[KEY_PROP]) return undefined
			const elementNew = Object.assign({}, element)
			delete elementNew[KEY_PROP]
			const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(elementNew)(
				this.getSPObject(element, clientContext)
			)
			return load(clientContext, spObject, opts)
		})
		if (isFilled(result)) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
	}

	async updateByQuery(opts) {
		const { result } = await this.iterator(async ({ element }) => {
			const { Folder, Query, Columns = {} } = element
			if (Query === undefined) throw new Error('Query is missed')
			const items = await this.of({ Folder, Query }).get({ ...opts, expanded: false })
			if (items.length) {
				const itemsToUpdate = items.map(item => {
					const columns = { ...Columns }
					columns[KEY_PROP] = item[KEY_PROP]
					return columns
				})
				return this.of(itemsToUpdate).update(opts)
			}
			return undefined
		})
		return result
	}

	async delete(opts = {}) {
		const { noRecycle, isSerial } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementID = element[KEY_PROP]
			if (!elementID) return undefined
			const spObject = this.getSPObject(element, clientContext)
			if (!spObject.isRoot) methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			return elementID
		})

		if (this.box.getCount()) {
			if (isSerial) {
				for (let i = 0; i < clientContexts.length; i += 1) {
					await executorJSOM(clientContexts[i])
				}
			} else {
				await Promise.all(clientContexts.map(executorJSOM))
			}
		}
		this.report(noRecycle ? 'delete' : 'recycle', opts)
		return prepareResponseJSOM(result, opts)
	}

	async deleteByQuery(opts) {
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

	async getDuplicates() {
		const items = await this
			.of({
				Scope: 'allItems',
				Limit: MAX_ITEMS_LIMIT
			})
			.get({
				view: ['ID', 'FSObjType', 'Modified', ...this.box.getIterable().map(el => el[KEY_PROP])]
			})

		const getHashedColumnName = itemData => {
			const values = []
			this.box.chain(column => {
				const value = itemData[column[KEY_PROP]]
				if (isExists(value)) values.push(value.get_lookupId ? value.get_lookupId() : value)
			})
			return MD5(values.join('.')).toString()
		}

		const itemsMap = items.reduce((map, item) => {
			const hashedColumnName = getHashedColumnName(item)
			if (!map.has(hashedColumnName)) map.set(hashedColumnName, [])
			map.get(hashedColumnName).push(item)
			return map
		}, new Map())

		const duplicatedsSorted = []
		for (const itemMap of itemsMap) {
			const [, duplicateds] = itemMap
			if (duplicateds.length > 1) {
				duplicatedsSorted.push(duplicateds.sort((a, b) => a.Modified.getTime() - b.Modified.getTime()))
			}
		}
		return duplicatedsSorted.map(arrayInit)
	}

	async deleteDuplicates(opts = {}) {
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

	async deleteEmpties(opts) {
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
						const itemToMerge = { ID: targetItem[KEY_PROP] }
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

	async erase(opts) {
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
		return prepareResponseJSOM(result, opts)
	}

	async createDiscussion(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const listSPObject = this.getListSPObject(this.listUrl, clientContext)
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
		await Promise.all(clientContexts.map(executorJSOM))
		this.report('create', opts)
		return prepareResponseJSOM(opts, result)
	}

	async createReply(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const { Columns } = element
			const spItemObject = this.getSPObject(element, clientContext)
			const spObject = SP.Utilities.Utility.createNewDiscussionReply(clientContext, spItemObject)
			const keys = Reflect.ownKeys(Columns)
			for (let i = 0; i < keys.length; i += 1) {
				const columnName = keys[i]
				spObject.set_item(columnName, Columns[columnName])
			}
			spObject.update()
			return spObject
		})
		await Promise.all(clientContexts.map(executorJSOM))
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	attachment(attachments) {
		return attachment(this, attachments)
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

	getSPObject(element, clientContext) {
		let camlQuery = new SP.CamlQuery()
		const { ID } = element
		const spListObject = this.getListSPObject(this.listUrl, clientContext)
		const contextUrl = clientContext.get_url()
		switch (typeOf(ID)) {
			case 'number': {
				const spObject = spListObject.getItemById(ID)
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
					const spObject = spListObject.getItemById(element.get_lookupId())
					return spObject
				} else {
					const { Folder, Page } = element
					if (Folder) {
						let fullUrl = ''
						const folderUrl = `${this.listFolder}${this.listUrl}/${Folder}`
						if (contextUrl === '/') {
							fullUrl = folderUrl
						} else {
							const isFullUrl = new RegExp(contextUrl, 'i').test(Folder)
							fullUrl = isFullUrl ? Folder : folderUrl
						}
						camlQuery.set_folderServerRelativeUrl(fullUrl)
					}
					camlQuery.set_viewXml(getCamlView(element))
					if (Page) {
						const { IsPrevious, ID, Columns } = Page
						const position = getInstanceEmpty(SP.ListItemCollectionPosition)
						position.set_pagingInfo(
							`${IsPrevious ? 'PagedPrev=TRUE&' : ''}Paged=TRUE&p_ID=${ID}${getPagingColumnsStr(Columns)}`
						)
						camlQuery.set_listItemCollectionPosition(position)
					}
				}
		}
		const spObject = spListObject.getItems(camlQuery)
		spObject.camlQuery = camlQuery
		return spObject
	}

	of(items) {
		return getInstance(this.constructor)(this.parent, items)
	}
}

export default getInstance(Item)
