/* eslint class-methods-use-this:0 */
/* eslint no-underscore-dangle:0 */
/* eslint prefer-destructuring:0 */
import {
	getClientContext,
	executeJSOM,
	prepareResponseJSOM,
	flatten,
	isArray,
	isNumberFilled,
	reduce,
	getInstance,
	isPropExists,
	isObject
} from '../lib/utility'

import { concatQueries } from '../lib/query-parser'

let defaultList
let customList
let customIdColumn
let customNameColumn = 'Title'
let customLoginColumn = 'Login'
let customEmailColumn = 'Email'
let customQuery

class User {
	constructor(parent, users, isSP) {
		this.name = 'user'
		this.parent = parent
		this.isArray = isArray(users)
		this.users = users ? (this.isArray ? flatten(users) : [users]) : []
		this.isSP = isSP
		this.isNative = this.isSP || (!customList || !customIdColumn)

		if (!defaultList) {
			defaultList = parent.of().list('User Information List')
		}
	}

	async get(opts) {
		if (this.users.length) {
			const el = this.users[0]

			if (isObject(el)) {
				return isPropExists(customIdColumn)(el)
					? this.getByUid(opts)
					: isPropExists(customLoginColumn)(el)
						? this.getByName(opts)
						: isPropExists(customEmailColumn)(el)
							? this.getByEMail(opts)
							: isPropExists(customLoginColumn)(el)
								? this.getByLogin(opts)
								: undefined
			}
			return el === '/'
				? this.getAll(opts)
				: isNumberFilled(el)
					? this.getByUid(opts)
					: /[а-яА-ЯЁё\s]/.test(el)
						? this.getByName(opts)
						: /\S+@\S+\.\S+/.test(el)
							? this.getByEMail(opts)
							: this.getByLogin(opts)
		}
		return this.getCurrent(opts)
	}

	async getCurrent(opts = {}) {
		if (this.isSP || !customList) {
			const clientContext = getClientContext('/')
			const spUser = clientContext.get_web().get_currentUser()
			return prepareResponseJSOM(await executeJSOM(clientContext, spUser, opts), opts)
		}
		const userId = window._spPageContextInfo
			? window._spPageContextInfo.userId
			: (await this.of(null, true).getCurrent({ view: 'Id' })).Id
		return (await customList.item(`Number ${customIdColumn} Eq ${userId}`).get(opts))[0]
	}

	async getByUid(opts = {}) {
		const userIds = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
				case 'Number':
					acc.push(item)
					break
				case 'Object': {
					const value = item[this.isNative ? 'ID' : customIdColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.FieldUserValue':
					acc.push(item.get_lookupId())
					break
				case 'SP.ListItem': {
					const userId = item.get_item(this.isNative ? 'ID' : customIdColumn)
					if (userId) acc.push(userId)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const query = `Number ${customIdColumn} In ${userIds}`
		const users = this.isNative
			? defaultList.item(userIds)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isArray || elements.length > 1 ? elements : elements[0]
	}

	async getByLogin(opts = {}) {
		let isByName = false
		const userLogins = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					if (/\|/.test(item)) isByName = true
					acc.push(item)
					break
				case 'Object': {
					const userName = item.UserName
					if (!userName) isByName = true
					acc.push(this.isNative
						? (userName || item.Name)
						: item.customLoginColumn)
					break
				}
				case 'SP.ListItem': {
					const userName = item.get_item('UserName')
					if (!userName) isByName = true
					acc.push(this.isNative
						? (userName || item.get_item('Name'))
						: item.get_item(customLoginColumn))
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const query = `${customLoginColumn} In ${userLogins}`
		const users = this.isNative
			? defaultList.item(`${isByName ? 'Name' : 'UserName'} In ${userLogins}`)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isArray || elements.length > 1 ? elements : elements[0]
	}

	async getByName(opts = {}) {
		const userNames = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object': {
					const value = item[this.isNative ? 'Title' : customNameColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.ListItem': {
					const title = item.get_item(this.isNative ? 'Title' : customNameColumn)
					if (title) acc.push(title)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const query = `${customNameColumn} Contains ${userNames}`
		return this.isNative
			? defaultList.item(`Title Contains ${userNames}`).get(opts)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query).get(opts)
	}

	async getByEMail(opts = {}) {
		const userEMails = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object': {
					const value = item[this.isNative ? 'EMail' : customEmailColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.ListItem': {
					const title = item.get_item(this.isNative ? 'EMail' : customEmailColumn)
					if (title) acc.push(title)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const query = `${customEmailColumn} In ${userEMails}`
		const users = this.isNative
			? defaultList.item(`EMail In ${userEMails}`)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isArray || elements.length > 1 ? elements : elements[0]
	}

	async getAll(opts) {
		return this.isNative ? defaultList.item().get(opts) : customList.item(customQuery).get(opts)
	}

	async create(opts) {
		if (!customList) throw new Error('Custom user list is missed')
		const usersToCreate = this.users.filter(el => el[customIdColumn] && el[customNameColumn])
		if (usersToCreate.length) {
			return customList.item(this.users).create(opts)
		}
		throw new Error('missing user id or Title')
	}

	async update(opts = {}) {
		if (this.isSP) {
			const results = await defaultList.item(this.users).update(opts)
			return this.isArray || results.length > 1 ? results : results[0]
		}
		if (!customList) throw new Error('Custom user list is missed')
		const ids = this.users.filter(el => !!el[customIdColumn])

		if (ids.length) {
			const userObjects = await this.of(ids).get({ view: ['ID', customIdColumn], mapBy: customIdColumn })
			const usersToUpdate = this.users.reduce((acc, el) => {
				const userID = userObjects[el[customIdColumn]] ? userObjects[el[customIdColumn]].ID : undefined
				if (userID) {
					const {
						[customIdColumn]: userId,
						...newEl
					} = el
					newEl.ID = userID
					acc.push(newEl)
				}
				return acc
			}, [])
			const results = await customList.item(usersToUpdate).update({ ...opts, view: ['ID', customIdColumn] })
			return this.isArray || results.length > 1 ? results : results[0]
		}
		throw new Error('missing user id')
	}

	setDefaults(params = {}) {
		if (params.customWebTitle && params.customListTitle) {
			customList = this.parent.of(params.customWebTitle).list(params.customListTitle)
		}
		if (params.defaultListTitle) defaultList = this.parent.of().list(params.defaultListTitle)
		if (params.customIdColumn) customIdColumn = params.customIdColumn
		if (params.customLoginColumn) customLoginColumn = params.customLoginColumn
		if (params.customNameColumn) customNameColumn = params.customNameColumn
		if (params.customEmailColumn) customEmailColumn = params.customEmailColumn
		if (params.customQuery) customQuery = params.customQuery
	}

	of(users, isSP) {
		return getInstance(this.constructor)(this.parent, users, isSP)
	}
}

export default getInstance(User)
