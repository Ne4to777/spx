import { getClientContext, prepareResponseJSOM, executeJSOM, methodEmpty, getInstance } from './../lib/utility'
import site from './../modules/web'
import recycleBin from './../modules/recycleBin'
import user from './../modules/user'
import tag from './../modules/tag'
import email from '../modules/email'

// Internal

const NAME = 'site'

const getSPObject = methodEmpty('get_site')

class Box {
	constructor(value) {
		this.value = { Url: value }
	}
	async chain(f) {
		return f(this.value)
	}
	join() {
		return this.value.Url
	}
}

const box = getInstance(Box)('/')

// Interface

site.tag = tag({ box, getSPObject })
site.user = user
site.recycleBin = recycleBin({ box, getSPObject })
site.email = email

site.get = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts)
	return prepareResponseJSOM(opts)(currentSPObjects)
}
site.getCustomListTemplates = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const web = clientContext.get_web()
	const templates = spObject.getCustomListTemplates(web)
	const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
	return prepareResponseJSOM(opts)(currentSPObjects)
}
site.getWebTemplates = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const templates = spObject.getWebTemplates(1033, 0)
	const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
	return prepareResponseJSOM(opts)(currentSPObjects)
}

site.time = {
	getCurrent: _ =>
		new Promise((resolve, reject) =>
			new SP.RequestExecutor('/').executeAsync({
				url: `/_api/web/RegionalSettings/TimeZone`,
				success: function(res) {
					resolve(new Date(res.headers.DATE))
				},
				error: reject
			})
		),
	getZone: async opts => {
		const clientContext = getClientContext('/')
		const result = await executeJSOM(clientContext)(
			clientContext
				.get_web()
				.get_regionalSettings()
				.get_timeZone()
		)(opts)
		return prepareResponseJSOM(opts)(result)
	}
}

export default site
