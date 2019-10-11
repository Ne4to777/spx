/* eslint-disable class-methods-use-this */
import {
	flatten,
	isArray,
	getInstance,
	getClientContext,
	executeJSOM,
	prepareResponseJSOM
} from '../lib/utility'

const KEY_PROP = 'Title'

class Group {
	constructor(parent, groups) {
		this.name = 'group'
		this.parent = parent
		this.isGroupsArray = isArray(groups)
		this.groups = groups ? (this.isGroupsArray ? flatten(groups) : [groups]) : []
	}

	async get(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const parentSPObject = this.parent.getSPObject(clientContext)
			const isCollection = hasUrlTailSlash(element.Url)
			const spObject = isCollection
				? this.getSPObjectCollection(parentSPObject)
				: this.getSPObject(element.Title, parentSPObject)
			return load(clientContext, spObject, opts)
		})

		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}


	async create(opts) {

	}

	async update(opts = {}) {

	}


	getSPObject(clientContext) {
		return clientContext.get_web()
	}


	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(clientContext).get_siteGroups()
	}

	of(groups) {
		return getInstance(this.constructor)(this.parent, groups)
	}
}

export default getInstance(Group)
