import {
	getClientContext,
	prepareResponseJSOM,
	ACTION_TYPES,
	executeJSOM,
} from './../lib/utility';

// Internal

const NAME = 'recycleBin';

const getSPObjectCollection = parent => clientContext => parent.getSPObject(clientContext).get_recycleBin();



// Interface

export default parent => {
	const report = actionType => ({ silent }) => !silent && console.log(`${ACTION_TYPES[actionType]} ${NAME} at ${parent.box.join()}`)
	return {
		get: async  opts => {
			const result = await parent.box.chain(async element => {
				const clientContext = getClientContext(element.Url);
				return executeJSOM(clientContext)(getSPObjectCollection(parent)(clientContext))(opts);
			})
			return prepareResponseJSOM(opts)(result)
		},
		restoreAll: async opts => {
			const result = await parent.box.chain(async element => {
				const clientContext = getClientContext(element.Url);
				const spObject = getSPObjectCollection(parent)(clientContext);
				spObject.restoreAll();
				return executeJSOM(clientContext)(spObject)(opts);;
			})
			report('restore')(opts);
			return prepareResponseJSOM(opts)(result)
		},
		deleteAll: async opts => {
			const result = await parent.box.chain(async element => {
				const clientContext = getClientContext(element.Url);
				const spObject = getSPObjectCollection(parent)(clientContext);
				spObject.deleteAll();
				return executeJSOM(clientContext)(spObject)(opts);;
			})
			report('delete')(opts);
			return prepareResponseJSOM(opts)(result)
		}
	}
}
