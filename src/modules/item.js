import * as utility from './../utility';
import * as cache from './../cache';
import spx from './../modules/site';
import {
	MD5
} from 'crypto-js';
import {
	getCamlQuery,
	getCamlView,
	concatQuery,
	camlLog
} from './../query_parser';

const _fieldsInfo = {};

export default class Item {
	constructor(parent, elementUrl) {
		this._parent = parent;
		this._contextUrlIsArray = this._parent._contextUrlIsArray;
		this._contextUrls = this._parent._contextUrls;
		this._listUrlIsArray = this._parent._elementUrlIsArray;
		this._listUrls = this._parent._elementUrls;
		this._elementUrl = elementUrl;
		this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
		this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl];
	}

	// Inteface

	async get(opts) { return this._execute(null, spObject => spObject, opts) }

	async create(opts) {
		await utility.initFieldsInfo(this._contextUrls, this._listUrls, this._fieldsInfo);
		return this._execute('create', async (spContextObject, listUrl, elementUrl) => {
			const elementUrlNew = Object.assign({}, elementUrl);
			const folder = elementUrl.folder;
			delete elementUrlNew.folder;
			const contextUrl = spContextObject.get_context().get_url();
			const itemCreationInfo = new SP.ListItemCreationInformation();
			if (folder) {
				const spxFolder = spx(contextUrl).list(listUrl).folder(folder);
				const folderOpts = { cached: true, silent: true };
				const folderUrl = await spxFolder.get(folderOpts).catch(async err => await spxFolder.create(folderOpts));
				itemCreationInfo.set_folderUrl(folderUrl.ServerRelativeUrl || folderUrl.FileRef);
			}
			const spObject = spContextObject.addItem(itemCreationInfo);
			utility.setItem(this._fieldsInfo[contextUrl][listUrl], spObject, elementUrlNew);
			return spObject;
		}, opts)
	}

	async update(opts) {
		await utility.initFieldsInfo(this._contextUrls, this._listUrls, this._fieldsInfo);
		return this._execute('update', async (spObject, listUrl, elementUrl) =>
			(utility.setItem(this._fieldsInfo[spObject.get_context().get_url()][listUrl], spObject, elementUrl), spObject), opts)
	}

	async delete(opts = {}) {
		return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject =>
			(spObject[opts.noRecycle ? 'deleteObject' : 'recycle'](), spObject), opts);
	}


	async getDuplicates(opts = {}) {
		let {
			columns = []
		} = opts;
		let items = await this.parent.item().get({
			view: ['ID', 'FSObjType', 'Modified', ...columns],
			scope: 'itemsAll',
			limit: 50000,
		})
		let itemsMap = new Map;
		let duplicatedsMap = new Map;
		let deleteMap = new Map;
		for (let item of items) {
			let hashedColumnName = getHashedColumnName(item);
			if (hashedColumnName !== void 0) {
				if (!itemsMap.has(hashedColumnName)) itemsMap.set(hashedColumnName, []);
				itemsMap.get(hashedColumnName).push(item);
			}
		}
		for (let [hashedColumnName, duplicateds] of itemsMap) {
			if (duplicateds.length > 1) {
				let itemToEval = duplicateds[0];
				for (let duplicated of duplicateds) {
					if (!duplicatedsMap.has(hashedColumnName)) duplicatedsMap.set(hashedColumnName, []);
					if (itemToEval.Modified.getTime() < duplicated.Modified.getTime()) {
						duplicatedsMap.get(hashedColumnName).push(itemToEval);
						deleteMap.set(itemToEval.ID, true);
						itemToEval = duplicated;
					} else {
						duplicatedsMap.get(hashedColumnName).push(duplicated);
						deleteMap.set(duplicated.ID, true);
					}
				}
			}
		}
		return [
			[...deleteMap.keys()], duplicatedsMap
		];

		function getHashedColumnName(itemData) {
			let values = [];
			if (typeOf(columns) === 'array') {
				for (let column of columns) {
					let value = itemData[column];
					if (value === null || value === void 0) {
						return
					} else {
						values.push(value.get_lookupId ? value.get_lookupId() : value);
					}
				}
			} else {
				value = itemData[columns];
				if (value === null || value === void 0) {
					return
				} else {
					values.push(value.get_lookupId ? value.get_lookupId() : value);
				}
			}
			return MD5(values.join('.')).toString()
		}
	}

	async deleteDuplicates(opts) {
		let [idsToDelete, itemsDuplicated] = await this.getDuplicates(opts);
		idsToDelete && idsToDelete.length && await this.parent.item(idsToDelete).delete();
	}

	async getEmpties(opts = {}) {
		let {
			columns
		} = opts;
		return this.parent.item(`${concatQuery(columns, 'isnull', 'or')}`).get({
			view: ['ID', ...columns],
			scope: 'itemsAll',
		})
	}

	async deleteEmpties(opts) {
		return this.parent.item(await this.getEmpties(opts)).delete();
	}

	async merge(opts = {}) {
		let {
			query,
			target,
			mediator = value => value,
			forced
		} = opts;
		let itemsToUpdate = [];
		let elements = typeOf(this.element) === 'array' ? this.element : [this.element];
		let items = await this.parent.item(query || `${concatQuery(elements, 'isnotnull', 'or')}`).get({
			...opts,
			view: ['ID', target, ...elements]
		});

		for (let item of items) {
			let row = {};
			for (let element of elements) {
				if (item[target] === null || forced) row[target] = await mediator(item[element])
			}
			if (Object.keys(row).length) {
				row.ID = item.ID;
				itemsToUpdate.push(row);
			}
		}
		return itemsToUpdate.length ? this.parent.item(itemsToUpdate).update(opts) : void 0;
	}
	async clear(opts = {}) {
		let itemsToClear = [];
		let elements = typeOf(this.element) === 'array' ? this.element : [this.element];
		let items = await this.parent.item(`${concatQuery(elements, 'isnotnull', 'or')}`).get({
			...opts,
			view: ['ID', ...elements],
		});
		for (let item of items) {
			let row = {};
			for (let element of elements) row[element] = null;
			if (Object.keys(row).length) {
				row.ID = item.ID;
				itemsToClear.push(row);
			}
		}
		return itemsToUpdate.length ? this.parent.item(itemsToClear).update(opts) : void 0;
	}

	// Internal

	get _name() { return 'item' }

	get _fieldsInfo() { return _fieldsInfo }

	async _execute(actionType, spObjectGetter, opts = {}) {
		let needToQuery;
		let isArrayCounter = 0;
		let { cached, showCaml, parallelized = actionType !== 'create' } = opts;
		const clientContexts = {};
		const spObjectsToCache = new Map;
		opts.view = opts.view || (actionType ? ['ID'] : void 0);
		const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
			let totalElements = 0;
			const spObjects = [];
			const contextUrls = contextUrl.split('/');
			let clientContext = utility.getClientContext(contextUrl);
			clientContexts[contextUrl] = [clientContext];

			for (let listUrl of this._listUrls) {
				for (let elementUrl of this._elementUrls) {
					if (!elementUrl) elementUrl = '';
					if (actionType && ++totalElements >= utility.REQUEST_BUNDLE_MAX_SIZE) {
						clientContext = utility.getClientContext(contextUrl);
						clientContexts[contextUrl].push(clientContext);
						totalElements = 0;
					}
					const spObject = await spObjectGetter(this._getSPObject(clientContext, listUrl, actionType === 'create' ? void 0 : elementUrl), listUrl, elementUrl);
					!!spObject.getEnumerator && isArrayCounter++;
					showCaml && spObject.camlQuery && camlLog(spObject.camlQuery.get_viewXml());
					const cachePath = MD5(spObject.get_path().$1_1.toString().match(/<Parameters>.*<\/Parameters>/) + (opts.view || '')).toString();
					const cachePaths = [...contextUrls, listUrl, this._name, cachePath];
					utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(cachePaths.slice(0, -1));
					if (actionType === 'delete' || actionType === 'recycle') {
						needToQuery = true;
					} else {
						const spObjectCached = cached ? cache.get(cachePaths) : null;
						if (cached && spObjectCached) {
							spObjects.push(spObjectCached);
						} else {
							needToQuery = true;
							const currentSPObjects = utility.load(clientContext, spObject, opts);
							!actionType && spObjectsToCache.set(cachePaths, currentSPObjects)
							spObjects.push(currentSPObjects);
						}
					}
				}
			}
			return spObjects;
		}))

		if (needToQuery) {
			await Promise.all(parallelized ?
				this._contextUrls.reduce((contextAcc, contextUrl) =>
					contextAcc.concat(clientContexts[contextUrl].map(clientContext => utility.executeQueryAsync(clientContext, opts))), []) :
				this._contextUrls.map(async (contextUrl) => {
					for (let clientContext of clientContexts[contextUrl]) await utility.executeQueryAsync(clientContext, opts)
				}));

			!actionType && spObjectsToCache.forEach((value, key) => cache.set(key, value))
		};


		this._log(actionType, opts);
		opts.isArray = isArrayCounter || this._contextUrlIsArray || this._listUrlIsArray || this._elementUrlIsArray;
		return utility.prepareResponseJSOM(elements, opts);
	}

	_getSPObject(clientContext, listUrl, elementUrl) {
		const list = this._parent._getSPObject(clientContext, listUrl);
		if (elementUrl === void 0) return list;
		let camlQuery = new SP.CamlQuery();
		switch (typeOf(elementUrl)) {
			case 'number':
				return list.getItemById(elementUrl);
			case 'string':
				camlQuery.set_viewXml(getCamlView(elementUrl));
				break;
			case 'object':
				if (elementUrl.set_viewXml) {
					camlQuery = elementUrl;
				} else if (elementUrl.get_lookupId) {
					return list.getItemById(elementUrl.get_lookupId());
				} else if (elementUrl.ID) {
					return list.getItemById(elementUrl.ID);
				} else {
					elementUrl.folder && camlQuery.set_folderServerRelativeUrl(`Lists/${listUrl}/${elementUrl.folder}`);
					camlQuery.set_viewXml(getCamlView(elementUrl.query, elementUrl));
					if (elementUrl.page) {
						const page = elementUrl.page;
						const pageStr = `${page.previous ? 'PagedPrev=TRUE&' : ''}Paged=TRUE&p_${page.field || 'ID'}=${page.value || ''}`;
						const position = new SP.ListItemCollectionPosition();
						position.set_pagingInfo(pageStr);
						camlQuery.set_listItemCollectionPosition(position);
					}
				}
				break;
		}
		const spObject = list.getItems(camlQuery);
		spObject.camlQuery = camlQuery;
		return spObject
	}

	_log(actionType, opts = {}) {
		!opts.silent && actionType &&
			console.log(`${
				utility.ACTION_TYPES[actionType]} ${
				this._contextUrls.length * this._listUrls.length * this._elementUrls.length} ${
				this._name}(s) at ${
				this._contextUrls.join(', ')} in ${
				this._listUrls.join(', ')}`);
	}
}