/* eslint-disable class-methods-use-this */
import { isNumber } from 'util'
import {
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
	webReport,
	getInstanceEmpty,
	setFields
} from '../lib/utility'

import user from './groupUser'

const KEY_PROP = 'Title'

const arrayValidator = pipe([removeEmptiesByProp(KEY_PROP), removeDuplicatedProp(KEY_PROP)])

const lifter = switchType({
	object: group => {
		const newGroup = Object.assign({}, group)
		const { Title, LoginName } = group
		newGroup[KEY_PROP] = LoginName || Title
		if (!newGroup.LoginName) newGroup.LoginName = newGroup[KEY_PROP]
		return newGroup
	},
	string: group => ({
		[KEY_PROP]: group,
		LoginName: group
	}),
	number: group => ({
		[KEY_PROP]: group,
		LoginName: group
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

class Group {
	constructor(parent, groups) {
		this.name = 'group'
		this.parent = parent
		this.box = getInstance(Box)(groups)
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
			const groupCreationInfo = getInstanceEmpty(SP.GroupCreationInformation)
			setFields({
				set_description: element.Description,
				set_title: element[KEY_PROP]
			})(groupCreationInfo)
			const newSPObject = spObject.add(groupCreationInfo)
			setFields({
				set_allowMembersEditMembership: element.AllowMembersEditMembership,
				set_allowRequestToJoinLeave: element.AllowRequestToJoinLeave,
				set_autoAcceptRequestToJoinLeave: element.AutoAcceptRequestToJoinLeave,
				set_onlyAllowMembersViewMembership: element.OnlyAllowMembersViewMembership,
				set_owner: element.Owner,
				set_requestToJoinLeaveEmailSetting: element.RequestToJoinLeaveEmailSetting
			})(newSPObject)
			newSPObject.update()
			return load(clientContext, newSPObject, opts)
		})

		if (this.count) await Promise.all(clientContexts.map(executorJSOM))
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	async update(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const spObject = this.getSPObject(element[KEY_PROP], clientContext)
			setFields({
				set_allowMembersEditMembership: element.AllowMembersEditMembership,
				set_allowRequestToJoinLeave: element.AllowRequestToJoinLeave,
				set_autoAcceptRequestToJoinLeave: element.AutoAcceptRequestToJoinLeave,
				set_description: element.Description,
				set_onlyAllowMembersViewMembership: element.OnlyAllowMembersViewMembership,
				set_owner: element.Owner,
				set_requestToJoinLeaveEmailSetting: element.RequestToJoinLeaveEmailSetting,
				set_title: element[KEY_PROP]
			})(spObject)
			spObject.update()
			return load(clientContext, spObject, opts)
		})

		if (this.count) await Promise.all(clientContexts.map(executorJSOM))
		this.report('update', opts)
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

	user(elements) {
		return user(this, elements)
	}


	getSPObject(element, clientContext) {
		return this.getSPObjectCollection(clientContext)[isNumber(element) ? 'getById' : 'getByName'](element)
	}


	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(clientContext).get_siteGroups()
	}

	report(actionType, opts = {}) {
		webReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			contextUrl: '/'
		})
	}

	of(groups) {
		return getInstance(this.constructor)(this.parent, groups)
	}
}

export default getInstance(Group)
