import * as utility from './../utility';
import * as cache from './../cache';

export default class RecycleBin {
	constructor(parent) {
		this._parent = parent;
		this._contextUrlIsArray = this._parent._contextUrlIsArray;
		this._contextUrls = this._parent._contextUrls;
	}

	// Inteface

	async get(opts) {
		return this._execute(null, spObject => (spObject.cachePath = 'propertiesCollection', spObject), opts);
	}
	async restoreAll(opts = {}) {
		return this._execute('restore', spObject =>
			(spObject.restoreAll(), spObject.cachePath = 'propertiesCollection', spObject), opts);
	}
	async deleteAll(opts = {}) {
		return this._execute('delete', spObject =>
			(spObject.deleteAll(), spObject.cachePath = 'propertiesCollection', spObject), opts);
	}

	// Internal

	get _name() { return 'recycleBin' }

	async _execute(actionType, spObjectGetter, opts = {}) {
		const { cached } = opts;
		const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
			const clientContext = utility.getClientContext(contextUrl);
			const spObject = spObjectGetter(this._getSPObject(clientContext));
			const contextSplits = contextUrl.split('/');
			const cachePaths = [...contextSplits, this._name, spObject.cachePath];
			utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(contextSplits);
			if (actionType) {
				await utility.executeQueryAsync(clientContext, opts);
			} else {
				const spObjectCached = cached ? cache.get(cachePaths) : null;
				if (cached && spObjectCached) {
					return spObjectCached;
				} else {
					const spObjects = utility.load(clientContext, spObject, opts);
					await utility.executeQueryAsync(clientContext, opts);
					cache.set(cachePaths, spObjects);
					return spObjects;
				}
			}
		}))

		this._log(actionType, opts);
		opts.isArray = true;
		return utility.prepareResponseJSOM(elements, opts);
	}

	_getSPObject(clientContext) { return this._parent._getSPObject(clientContext).get_recycleBin() }

	_log(actionType, opts = {}) {
		!opts.silent && actionType &&
			console.log(`${
				utility.ACTION_TYPES[actionType]} ${
				this._name} at ${
				this._contextUrls.join(', ')}`);
	}
}