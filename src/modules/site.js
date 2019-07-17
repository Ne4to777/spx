import {
	getClientContext, prepareResponseJSOM, executeJSOM, methodEmpty, getInstance
} from '../lib/utility'
// import _web from './web'
import _recycleBin from './recycleBin'
// import _user from './user'
// import _tag from './tag'
// import _email from './email'
// import _time from './time'

// Internal

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

// export const web = _web
// export const tag = _tag({ box, getSPObject })
// export const user = _user
export const recycleBin = _recycleBin({ box, getSPObject })
// export const email = _email

// export const time = _time

export const getSite = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts)
	console.log(currentSPObjects)

	return prepareResponseJSOM(opts)(currentSPObjects)
}
export const getCustomListTemplates = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const templates = spObject.getCustomListTemplates(clientContext.get_web())
	const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
	return prepareResponseJSOM(opts)(currentSPObjects)
}
export const getWebTemplates = async opts => {
	const clientContext = getClientContext('/')
	const spObject = getSPObject(clientContext)
	const templates = spObject.getWebTemplates(1033, 0)
	const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
	return prepareResponseJSOM(opts)(currentSPObjects)
}
