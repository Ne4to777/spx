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
	constructor(parent, users) {
		this.name = 'user'
		this.parent = parent
		this.isUsersArray = isArray(users)
		this.users = users ? (this.isUsersArray ? flatten(users) : [users]) : []

		if (!defaultList) {
			defaultList = parent.of().list('User Information List')
		}
	}

	async	get(opts = {}) {
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
		const { isSP } = opts
		if (isSP || !customList) {
			const clientContext = getClientContext('/')
			const spUser = clientContext.get_web().get_currentUser()
			return prepareResponseJSOM(await executeJSOM(clientContext, spUser, opts), opts)
		}
		const userId = window._spPageContextInfo
			? window._spPageContextInfo.userId
			: (await this.getCurrent({ view: 'Id', isSP: true })).Id
		return (await customList.item(`Number ${customIdColumn} Eq ${userId}`).get(opts))[0]
	}

	async getByUid(opts = {}) {
		const { isSP } = opts
		const isNative = isSP || (!customList || !customIdColumn)
		const userIds = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
				case 'Number':
					acc.push(item)
					break
				case 'Object': {
					const value = item[isNative ? 'ID' : customIdColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.FieldUserValue':
					acc.push(item.get_lookupId())
					break
				case 'SP.ListItem': {
					const userId = item.get_item(isNative ? 'ID' : customIdColumn)
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
		const users = isNative
			? defaultList.item(userIds)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async getByLogin(opts = {}) {
		const { isSP } = opts
		const isNative = isSP || !customList
		const userLogins = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object': {
					const value = item[isNative ? 'UserName' : customLoginColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.ListItem': {
					const login = item.get_item(isNative ? 'UserName' : customLoginColumn)
					if (login) acc.push(login)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const query = `${customLoginColumn} In ${userLogins}`
		const users = isNative
			? defaultList.item(`UserName In ${userLogins}`)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async getByName(opts = {}) {
		const { isSP } = opts
		const isNative = isSP || !customList
		const userNames = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object': {
					const value = item[isNative ? 'Title' : customNameColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.ListItem': {
					const title = item.get_item(isNative ? 'Title' : customNameColumn)
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
		return isNative
			? defaultList.item(`Title Contains ${userNames}`).get(opts)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query).get(opts)
	}

	async getByEMail(opts = {}) {
		const { isSP } = opts
		const isNative = isSP || !customList
		const userEMails = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object': {
					const value = item[isNative ? 'EMail' : customEmailColumn]
					if (value) acc.push(value)
					break
				}
				case 'SP.ListItem': {
					const title = item.get_item(isNative ? 'EMail' : customEmailColumn)
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
		const users = isNative
			? defaultList.item(`EMail In ${userEMails}`)
			: customList.item(customQuery ? concatQueries('and')([customQuery, query]) : query)
		const elements = await users.get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async getAll(opts = {}) {
		return opts.isSP || !customList
			? defaultList.item().get(opts)
			: customList.item(customQuery).get(opts)
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
		const { isSP } = opts
		if (isSP) {
			const usersToUpdate = this.users.reduce((acc, el) => {
				if (el[customIdColumn]) {
					const {
						[customIdColumn]: userId,
						...newEl
					} = el
					newEl.ID = userId
					acc.push(newEl)
				}
				return acc
			}, [])
			const results = await defaultList
				.item(usersToUpdate)
				.update(opts)
			return this.isUsersArray || results.length > 1 ? results : results[0]
		}
		if (!customList) throw new Error('Custom user list is missed')
		const ids = this.users.filter(el => !!el[customIdColumn])
		if (ids.length) {
			const userObjects = await this
				.of(ids)
				.get({ view: ['ID', customIdColumn], mapBy: customIdColumn })
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
			return this.isUsersArray || results.length > 1 ? results : results[0]
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

	of(users) {
		return getInstance(this.constructor)(this.parent, users)
	}
}

export default getInstance(User)
