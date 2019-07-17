import {
	getClientContext, prepareResponseJSOM, ACTION_TYPES, executeJSOM, getInstance
} from '../lib/utility'

const NAME = 'recycleBin'

class RecycleBin {
	constructor(parent) {
		this.parent = parent
	}

	async get(opts) {
		const result = await this.parent.box.chain(async element => {
			const clientContext = getClientContext(element.Url)
			return executeJSOM(clientContext)(this.getSPObjectCollection(clientContext))(opts)
		})
		return prepareResponseJSOM(opts)(result)
	}

	async	restoreAll(opts) {
		const result = await this.parent.box.chain(async element => {
			const clientContext = getClientContext(element.Url)
			const spObject = this.getSPObjectCollection(clientContext)
			spObject.restoreAll()
			return executeJSOM(clientContext)(spObject)(opts)
		})
		this.report('restore')(opts)
		return prepareResponseJSOM(opts)(result)
	}

	async deleteAll(opts) {
		const result = await this.parent.box.chain(async element => {
			const clientContext = getClientContext(element.Url)
			const spObject = this.getSPObjectCollection(clientContext)
			spObject.deleteAll()
			return executeJSOM(clientContext)(spObject)(opts)
		})
		this.report('delete')(opts)
		return prepareResponseJSOM(opts)(result)
	}

	report(actionType, opts = {}) {
		if (!opts.silent) console.log(`${ACTION_TYPES[actionType]} ${NAME} at ${this.parent.box.join()}`)
	}

	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(clientContext).get_recycleBin()
	}
}

export default getInstance(RecycleBin)
