import {
	getClientContext,
	prepareResponseJSOM,
	ACTION_TYPES,
	executeJSOM,
	getInstance,
	report
} from '../lib/utility'

class RecycleBin {
	constructor(parent) {
		this.name = 'recycleBin'
		this.contextUrl = parent.box.head().Url
		this.getSPObject = this.contextUrl ? parent.getSPObject : parent.getSiteSPObject
	}

	async get(opts) {
		const clientContext = getClientContext(this.contextUrl)
		const result = await executeJSOM(clientContext)(this.getSPObjectCollection(clientContext))(opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	restoreAll(opts) {
		const clientContext = getClientContext(this.contextUrl)
		const spObject = this.getSPObjectCollection(clientContext)
		spObject.restoreAll()
		const result = await executeJSOM(clientContext)(spObject)(opts)
		this.report('restore', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async deleteAll(opts) {
		const clientContext = getClientContext(this.contextUrl)
		const spObject = this.getSPObjectCollection(clientContext)
		spObject.deleteAll()
		const result = await executeJSOM(clientContext)(spObject)(opts)
		this.report('delete', opts)
		return prepareResponseJSOM(opts)(result)
	}

	report(actionType, opts = {}) {
		report(`${ACTION_TYPES[actionType]} ${this.name} at ${this.contextUrl}`, opts)
	}

	getSPObjectCollection(clientContext) {
		return this.getSPObject(clientContext).get_recycleBin()
	}
}

export default getInstance(RecycleBin)
