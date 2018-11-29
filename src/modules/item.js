import * as utility from './../utility';
import * as cache from './../cache';
import spx from './../modules/site';
import {
	MD5
} from 'crypto-js';
import {
	getCamlQuery,
	getCamlView,
	joinQuery,
	camlLog,
	concatQuery
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

	async updateByQuery(opts = {}) {
		return utility.prepareResponseJSOM(await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
				listAcc.concat(this._elementUrls.map(async  el => {
					const { query, columns = {} } = el;
					if (query === void 0) throw new Error('query is missed');
					const list = spx(contextUrl).list(listUrl);
					const items = await list.item(query).get(opts);
					if (items.length) {
						const itemsToUpdate = items.map(item => {
							const row = Object.assign({}, columns);
							row.ID = item.ID;
							return row
						});
						return list.item(itemsToUpdate).update({ expanded: true });
					}
				})), [])), [])))
	}

	async delete(opts = {}) {
		return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject =>
			(spObject[opts.noRecycle ? 'deleteObject' : 'recycle'](), spObject), opts);
	}

	async deleteByQuery(opts = {}) {
		return utility.prepareResponseJSOM(await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
				listAcc.concat(this._elementUrls.map(async  el => {
					const list = spx(contextUrl).list(listUrl);
					const items = await list.item(el).get({ view: ['ID'], ...opts });
					if (items.length) {
						const ids = items.map(item => item.ID);
						return list.item(ids).delete(opts);
					}
				})), [])), [])), { expanded: true })
	}


	async getDuplicates() {
		return this._operateDuplicates()
	}

	async deleteDuplicates() {
		return this._operateDuplicates({ delete: true })
	}

	async getEmpties(opts = {}) {
		return this._parent.item(`${joinQuery(this._elementUrls, 'or', 'isnull')}`).get({ ...opts, scope: 'itemsAll', limit: utility.MAX_ITEMS_LIMIT })
	}

	async deleteEmpties(opts) {
		return this._parent.item((await this.getEmpties(opts)).map(el => el.ID)).delete();
	}

	async merge(opts = {}) {
		return utility.prepareResponseJSOM(await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
				listAcc.concat(this._elementUrls.map(async  elementUrl => {
					const { key, forced, target, source, mediator = x => x } = elementUrl;
					const sourceColumn = source.column;
					const sourceQuery = source.query;
					const sourceWeb = source.web;
					const sourceList = source.list;
					const targetColumn = target.column;
					const targetQuery = target.query;

					const sourceItems = await spx(sourceWeb).list(sourceList).item(concatQuery(`${key} IsNotNull`, sourceQuery)).get({
						view: ['ID', key, sourceColumn],
						scope: 'itemsAll',
						limit: utility.MAX_ITEMS_LIMIT,
						...opts
					});
					let targetList = spx(contextUrl).list(listUrl);
					const targetItems = await targetList.item(concatQuery(`${key} IsNotNull`, targetQuery)).get({
						view: ['ID', key, targetColumn],
						scope: 'itemsAll',
						limit: utility.MAX_ITEMS_LIMIT,
						groupBy: key,
						...opts
					})

					const itemsToMerge = [];
					for (let sourceItem of sourceItems) {
						const targetItem = targetItems[sourceItem[key]];
						if (targetItem) {
							const itemToMerge = { ID: targetItem.ID };
							if (targetItem[targetColumn] === null || forced) {
								itemToMerge[targetColumn] = await mediator(sourceItem[sourceColumn], sourceItem);
								itemsToMerge.push(itemToMerge);
							}
						}
					}
					return targetList.item(itemsToMerge).update(opts);
				})), [])), [])), { expanded: true })
	}

	async clear(opts) {
		return utility.prepareResponseJSOM(await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
				listAcc.concat(this._elementUrls.map(async  el => {
					const { query = '', columns = {} } = el;
					const list = spx(contextUrl).list(listUrl);
					const items = await list.item(query).get(opts);
					if (items.length) {
						const itemsToUpdate = items.map(item => {
							const row = {};
							columns.map(col => row[col] = null)
							row.ID = item.ID;
							return row
						});
						return list.item(itemsToUpdate).update({ expanded: true });
					}
				})), [])), [])))
	}

	// Internal

	get _name() { return 'item' }

	get _fieldsInfo() { return _fieldsInfo }

	async _execute(actionType, spObjectGetter, opts = {}) {
		let needToQuery;
		let isArrayCounter = 0;
		const clientContexts = {};
		const spObjectsToCache = new Map;
		let { cached, showCaml, parallelized = actionType !== 'create' } = opts;
		opts.view = opts.view || (actionType ? ['ID'] : void 0);
		const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) => {
			let totalElements = 0;
			const contextUrls = contextUrl.split('/');
			let clientContext = utility.getClientContext(contextUrl);
			clientContexts[contextUrl] = [clientContext];
			return contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
				listAcc.concat(this._elementUrls.map(async elementUrl => {
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
							return spObjectCached;
						} else {
							needToQuery = true;
							const currentSPObjects = utility.load(clientContext, spObject, opts);
							!actionType && spObjectsToCache.set(cachePaths, currentSPObjects)
							return currentSPObjects;
						}
					}
				})), []))
		}, []))

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

	async _operateClear() {
		const items = await this._parent.item(`${joinQuery(this._elementUrls, 'or', 'isnull')}`).get({
			...opts,
			view: ['ID', ...this._elementUrls],
		});
		const itemsToClear = items.reduce((acc, item) => {
			const row = {};
			for (let element of this._elementUrls) row[element] = null;
			if (Object.keys(row).length) {
				row.ID = item.ID;
				acc.push(row);
			}
			return acc;
		}, []);
		return itemsToClear.length ? this._parent.item(itemsToClear).update(opts) : void 0;
	}

	async _operateDuplicates(params = {}) {
		const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._listUrls.map(async (listUrl) => {
				const items = await spx(contextUrl).list(listUrl).item().get({
					view: ['ID', 'FSObjType', 'Modified', ...this._elementUrl],
					scope: 'itemsAll',
					limit: utility.MAX_ITEMS_LIMIT,
				});

				const itemsMap = new Map;
				const duplicatedsMap = new Map;
				const deleteMap = new Map;
				const getHashedColumnName = itemData => {
					const values = [];
					for (let column of this._elementUrls) {
						const value = itemData[column];
						if (value === null || value === void 0) {
							return
						} else {
							values.push(value.get_lookupId ? value.get_lookupId() : value);
						}
					}
					return MD5(values.join('.')).toString()
				}
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
				if (params.delete) {
					await spx(contextUrl).list(listUrl).item([...deleteMap.keys()]).delete();
				} else {
					let array = [];
					for (let [k, v] of duplicatedsMap) array.push(v);
					return array
				}
			})), []))
		return elements;
	}
}