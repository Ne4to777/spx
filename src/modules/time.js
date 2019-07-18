import {
	getClientContext, prepareResponseJSOM, executeJSOM
} from '../lib/utility'

export const getCurrent = () => new Promise(
	(resolve, reject) => new SP.RequestExecutor('/').executeAsync({
		url: '/_api/web/RegionalSettings/TimeZone',
		success(res) {
			resolve(new Date(res.headers.DATE))
		},
		error: reject
	})
)
export const getZone = async opts => {
	const clientContext = getClientContext('/')
	const result = await executeJSOM(clientContext)(clientContext
		.get_web()
		.get_regionalSettings()
		.get_timeZone())(opts)
	return prepareResponseJSOM(opts)(result)
}
