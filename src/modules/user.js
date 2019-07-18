import {
	getClientContext,
	executeJSOM,
	prop,
	prepareResponseJSOM,
	flatten,
	isArray,
	isNumber,
	typeOf,
	reduce,
	getInstance
} from '../lib/utility'

import web from './web'

// let defaultUsersList = web().list('User Information List')
let customUsersList

const setCustomUsersList = (data = {}) => {
	if (data.listTitle) {
		customUsersList = web(data.webTitle).list(data.listTitle)
	} else {
		throw new Error('Wrong data object. Need {webTitle, listTitle}')
	}
}

const setDefaultUsersList = title => {
	defaultUsersList = web().list(title)
}

const get = async (opts = {}) => {
	const { isSP } = opts
	if (isSP || !customUsersList) {
		const clientContext = getClientContext('/')
		const spUser = clientContext.get_web().get_currentUser()
		return prepareResponseJSOM(opts)(await executeJSOM(clientContext)(spUser)(opts))
	}
	const uid = _spPageContextInfo
		? _spPageContextInfo.userId
		: (await user.get({ view: 'Id', isSP: true })).Id
	return (await customUsersList.item(`Number uid Eq ${uid}`).get(opts))[0]
}

const getAll = async (opts = {}) => opts.isSP || !customUsersList ? defaultUsersList.item() : customUsersList.item(
	'Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))'
).get(opts)

const deleteWithMissedUid = async opts => {
	if (!customUsersList) throw new Error('Custom user list is missed')
	const userObjects = await customUsersList.item('Number uid IsNull').get(opts)
	return customUsersList.item(userObjects.map(prop('ID'))).delete(opts)
}

// Interface

class User {
	constructor(users) {
		this.isUsersArray = isArray(users)
		this.users = users ? (this.isUsersArray ? flatten(users) : [users]) : []
	}

	async	get(opts = {}) {
		if (this.users.length) {
			const el = this.users[0]
			const values = this.users
			return isNumber(el) || Number(el) === el || (typeOf(el) === 'object' && el.uid)
				? this.getByUid(values, opts)
				: /\s/.test(el) || /[а-яА-ЯЁё]/.test(el)
					? this.getByName(values)(opts)
					: /@.+\./.test(el)
						? this.getByEMail(values, opts)
						: this.getByLogin(values, opts)
		}
		return getAll(opts)
	}

	async	create(opts) {
		if (!customUsersList) throw new Error('Custom user list is missed')
		const usersToCreate = this.users.filter(el => el.uid && el.Title)
		return usersToCreate.length
			? customUsersList.item(this.users).create(opts)
			: new Promise((resolve, reject) => reject(new Error('missing uid or Title')))
	}

	async getByUid(opts = {}) {
		const userIds = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
				case 'Number':
					acc.push(item)
					break
				case 'Object':
					if (item.uid) acc.push(item.uid)
					break
				case 'SP.User':
					acc.push(item.get_id())
					break
				case 'SP.FieldUserValue':
					acc.push(item.get_lookupId())
					break
				case 'SP.ListItem': {
					const uid = item.get_item('uid')
					if (uid) acc.push(uid)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const users = opts.isSP || !customUsersList
			? defaultUsersList.item(userIds)
			: customUsersList.item(`Number uid In ${userIds}`)
		const elements = await users.get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async getByLogin(opts = {}) {
		const userLogins = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object':
					if (item.Login) acc.push(item.Login)
					break
				case 'SP.User':
					acc.push(item.get_loginName())
					break
				case 'SP.ListItem': {
					const login = item.get_item('Login')
					if (login) acc.push(login)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const users = opts.isSP || !customUsersList
			? defaultUsersList.item(`UserName In ${userLogins}`)
			: customUsersList.item(`Login In ${userLogins}`)
		const elements = await users.get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async getByName(opts = {}) {
		const userNames = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object':
					if (item.Title) acc.push(item.Title)
					break
				case 'SP.User':
					acc.push(item.get_title())
					break
				case 'SP.ListItem': {
					const title = item.get_item('Title')
					if (title) acc.push(title)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const list = opts.isSP || !customUsersList ? defaultUsersList : customUsersList
		return list.item(`Title BeginsWith ${userNames}`).get(opts)
	}

	async getByEMail(opts = {}) {
		const userEMails = reduce(acc => item => {
			switch (item.constructor.getName()) {
				case 'String':
					acc.push(item)
					break
				case 'Object':
					if (item.EMail) acc.push(item.EMail)
					break
				case 'SP.User':
					acc.push(item.get_email())
					break
				case 'SP.ListItem': {
					const title = item.get_item('EMail')
					if (title) acc.push(title)
					break
				}
				default: {
					// default
				}
			}
			return acc
		})([])(this.users)
		const list = opts.isSP || !customUsersList ? defaultUsersList : customUsersList
		const elements = await list.item(`EMail In ${userEMails}`).get(opts)
		return this.isUsersArray || elements.length > 1 ? elements : elements[0]
	}

	async	update(opts = {}) {
		const { isSP } = opts
		if (isSP) {
			const usersToUpdate = this.users.reduce((acc, el) => {
				if (el.uid) {
					const {
						uid,
						...newEl
					} = el
					newEl.ID = uid
					acc.push(newEl)
				}
				return acc
			}, [])
			const results = await defaultUsersList
				.item(usersToUpdate)
				.update(opts)
			return this.isUsersArray || results.length > 1 ? results : results[0]
		}
		if (!customUsersList) throw new Error('Custom user list is missed')
		const ids = this.users.filter(el => !!el.uid)
		if (ids.length) {
			const userObjects = await web.user(ids).get({ view: ['ID', 'uid'], groupBy: 'uid' })
			const usersToUpdate = this.users.reduce((acc, el) => {
				const userID = userObjects[el.uid] ? userObjects[el.uid].ID : undefined
				if (userID) {
					const {
						uid,
						...newEl
					} = el
					newEl.ID = userID
					acc.push(el)
				}
				return acc
			}, [])
			const results = await customUsersList.item(usersToUpdate).update({ ...opts, view: ['ID', 'uid'] })
			return this.isUsersArray || results.length > 1 ? results : results[0]
		}
		return new Promise((resolve, reject) => reject(new Error('missing uid')))
	}
}

const user = getInstance(User)

user.get = get
user.getAll = getAll
user.deleteWithMissedUid = deleteWithMissedUid
user.setDefaultUsersList = setDefaultUsersList
user.setCustomUsersList = setCustomUsersList

export default user
