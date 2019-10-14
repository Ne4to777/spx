/* eslint-disable class-methods-use-this */
import { isNumber } from 'util'
import {
	isArray,
	getInstance,
	executorJSOM,
	load,
	AbstractBox,
	deep1Iterator,
	prepareResponseJSOM,
	removeEmptiesByProp,
	removeDuplicatedProp,
	switchType,
	pipe,
	setFields,
	getInstanceEmpty,
	listReport
} from '../lib/utility'


const KEY_PROP = 'Title'

const arrayValidator = pipe([removeEmptiesByProp(KEY_PROP), removeDuplicatedProp(KEY_PROP)])

const lifter = switchType({
	object: user => {
		const newGroup = Object.assign({}, user)
		const { Title, LoginName } = user
		newGroup[KEY_PROP] = LoginName || Title
		if (!newGroup.LoginName) newGroup.LoginName = newGroup[KEY_PROP]
		return newGroup
	},
	string: user => ({
		[KEY_PROP]: user,
		LoginName: user
	}),
	number: user => ({
		[KEY_PROP]: user,
		LoginName: user
	}),
	default: () => ({
		[KEY_PROP]: undefined,
		LoginName: undefined
	})
})

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter, arrayValidator)
		this.prop = KEY_PROP
		this.joinProp = KEY_PROP
	}
}

class GroupUser {
	constructor(parent, groupUsers) {
		this.name = 'user'
		this.parent = parent
		this.isArray = isArray(groupUsers)
		this.box = getInstance(Box)(groupUsers)
		this.groupName = this.parent.box.getHead()[KEY_PROP]
		this.count = this.box.getCount()
		this.iterator = deep1Iterator({
			elementBox: this.box
		})
	}

	async get(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element.Title
			const spObject = title
				? this.getSPObject(title, clientContext)
				: this.getSPObjectCollection(clientContext)
			return load(clientContext, spObject, opts)
		})

		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}

	async create(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const spObject = this.getSPObjectCollection(clientContext)
			const userCreationInfo = getInstanceEmpty(SP.UserCreationInformation)
			setFields({
				set_email: element.Email,
				set_loginName: element.LoginName,
				set_title: element[KEY_PROP]
			})(userCreationInfo)
			const newSPObject = spObject.add(userCreationInfo)
			return load(clientContext, newSPObject, opts)
		})

		if (this.count) await Promise.all(clientContexts.map(executorJSOM))
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}


	async delete(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			const spObject = this.getSPObjectCollection(clientContext)
			spObject[isNumber(title) ? 'removeById' : 'removeByLoginName'](title)
			return title
		})

		if (this.count) await Promise.all(clientContexts.map(executorJSOM))
		this.report('delete', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(element, clientContext) {
		return this.getSPObjectCollection(clientContext)[
			isNumber(element)
				? 'getById'
				: /@.+\./.test(element)
					? 'getByEmail'
					: 'getByLoginName'
		](element)
	}


	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(this.groupName, clientContext).get_users()
	}

	report(actionType, opts = {}) {
		listReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			contextUrl: '/',
			listUrl: this.groupName
		})
	}

	of(groupUsers) {
		return getInstance(this.constructor)(this.parent, groupUsers)
	}
}

export default getInstance(GroupUser)
