import {
	getClientContext,
	prepareResponseJSOM,
	ACTION_TYPES,
	executeJSOM,
} from './../lib/utility';

// Internal

const NAME = 'recycleBin';

const getSPObjectCollection = parent => clientContext => parent.getSPObject(clientContext).get_recycleBin();

const report = ({ silent, actionType, box }) =>
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME} at: ${box.join()}`)

// Interface

export default parent => {
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
			report({ ...opts, actionType: 'restore', box: parent.box })
			return prepareResponseJSOM(opts)(result)
		},
		deleteAll: async opts => {
			const result = await parent.box.chain(async element => {
				const clientContext = getClientContext(element.Url);
				const spObject = getSPObjectCollection(parent)(clientContext);
				spObject.deleteAll();
				return executeJSOM(clientContext)(spObject)(opts);;
			})
			report({ ...opts, actionType: 'delete', box: parent.box })
			return prepareResponseJSOM(opts)(result)
		}
	}
} 
