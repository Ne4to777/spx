import {
	ACTION_TYPES_TO_UNSET,
	getClientContext,
	prepareResponseJSOM,
	ACTION_TYPES,
	urlSplit,
	executorJSOM,
	executeJSOM,
	overstep,
	methodEmpty,
	identity
} from './../utility';
import * as cache from './../cache';

// Internal

const NAME = 'recycleBin';

const getSPObjectCollection = parent => clientContext => parent.getSPObject(clientContext).get_recycleBin();

const report = ({ silent, actionType }) => parent => spObjects => (
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${parent.box.joinUrls()}`),
	spObjects
)

const execute = parent => cacheLeaf => actionType => spObjectGetter => (opts = {}) => parent.box
	.chainAsync(async element => {
		const { cached } = opts;
		const contextUrl = element.Url;
		const clientContext = getClientContext(contextUrl);
		const contextUrlSplits = urlSplit(contextUrl);
		const spObject = spObjectGetter(getSPObjectCollection(parent)(clientContext));
		const cachePath = [...contextUrlSplits, NAME, cacheLeaf + 'Collection'];
		ACTION_TYPES_TO_UNSET[actionType] && cache.unset(contextUrlSplits);
		if (actionType === 'delete' || actionType === 'recycle') return executorJSOM(clientContext);
		const spObjectCached = cached ? cache.get(cachePath) : null;
		if (cached && spObjectCached) return spObjectCached;
		const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts);
		cache.set(currentSPObjects)(cachePath)
		return currentSPObjects;
	})
	.then(report({ ...opts, actionType })(parent))
	.then(prepareResponseJSOM(opts))


// Interface

export default parent => {
	const executeBinded = execute(parent)('properties');
	return {
		get: executeBinded(null)(identity),
		restoreAll: executeBinded('restore')(overstep(methodEmpty('restoreAll'))),
		deleteAll: executeBinded('delete')(overstep(methodEmpty('deleteAll')))
	}
} 
