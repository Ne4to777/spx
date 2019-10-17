/* eslint class-methods-use-this:0 */
import {
	getClientContext,
	prepareResponseJSOM,
	executeJSOM,
	getInstance
} from '../lib/utility'

class Time {
	constructor(parent) {
		this.name = 'time'
		this.parent = parent
		this.contextUrl = this.parent.contextUrl || '/'
	}

	async get() {
		return new Promise(
			(resolve, reject) => new SP.RequestExecutor(this.contextUrl).executeAsync({
				url: '/_api/web/RegionalSettings/TimeZone',
				success(res) {
					resolve(new Date(res.headers.DATE))
				},
				error: reject
			})
		)
	}

	async getZone(opts) {
		const clientContext = getClientContext(this.contextUrl)
		const spObject = this.parent.getSPObject(clientContext).get_regionalSettings().get_timeZone()
		const result = await executeJSOM(clientContext, spObject, opts)
		return prepareResponseJSOM(result, opts)
	}
}

export default getInstance(Time)
