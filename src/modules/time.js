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
	}

	async	get() {
		return new Promise(
			(resolve, reject) => new SP.RequestExecutor('/').executeAsync({
				url: '/_api/web/RegionalSettings/TimeZone',
				success(res) {
					resolve(new Date(res.headers.DATE))
				},
				error: reject
			})
		)
	}

	async	getZone(opts) {
		const clientContext = getClientContext('/')
		const result = await executeJSOM(clientContext)(clientContext
			.get_web()
			.get_regionalSettings()
			.get_timeZone())(opts)
		return prepareResponseJSOM(opts)(result)
	}
}

export default getInstance(Time)
