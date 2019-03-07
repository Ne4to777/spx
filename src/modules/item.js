import {
	ACTION_TYPES_TO_UNSET,
	REQUEST_BUNDLE_MAX_SIZE,
	ACTION_TYPES,
	MAX_ITEMS_LIMIT,
	IS_DELETE_ACTION,
	Box,
	getInstance,
	prepareResponseJSOM,
	getClientContext,
	urlSplit,
	load,
	executorJSOM,
	getInstanceEmpty,
	slice,
	setItem,
	getColumns,
	prop,
	overstep,
	methodEmpty,
	identity,
	arrayInit,
	isNumber,
	typeOf
} from './../utility';
import * as cache from './../cache';
import {
	MD5
} from 'crypto-js';
import {
	getCamlView,
	joinQueries,
	camlLog,
	concatQueries
} from './../query_parser';
import site from './../modules/site';

// Internal

const NAME = 'item';

const getSPObject = element => parentElement => {
	let camlQuery = new SP.CamlQuery();
	const ID = element.ID;
	switch (typeOf(ID)) {
		case 'number': return parentElement.getItemById(ID);
		case 'string': camlQuery.set_viewXml(getCamlView()(ID)); break;
		default:
			if (element.set_viewXml) {
				camlQuery = element;
			} else if (element.get_lookupId) {
				return parentElement.getItemById(element.get_lookupId());
			} else {
				const { Folder, Query, Page, IsLibrary } = element;
				if (Folder) {
					const isFullUrl = new RegExp(parentElement.get_context().get_url(), 'i').test(Folder);
					camlQuery.set_folderServerRelativeUrl(isFullUrl ? Folder : `${IsLibrary ? '' : 'Lists/'}${parentElement.listUrl}/${Folder}`);
				}
				element && camlQuery.set_viewXml(getCamlView(element)(Query));
				if (Page) {
					const position = new SP.ListItemCollectionPosition();
					const { IsPrevious, Id, Column, Value } = Page;
					position.set_pagingInfo(`${IsPrevious ? 'PagedPrev=TRUE&' : ''}Paged=TRUE&p_ID=${Id}&p_${Column || 'ID'}=${Value || ''}`);
					camlQuery.set_listItemCollectionPosition(position);
				}
			}
	}
	const spObject = parentElement.getItems(camlQuery);
	spObject.camlQuery = camlQuery;
	spObject.listUrl = parentElement.listUrl;
	return spObject
}

const listIterator = it => f => it.parent.parent.box.chainAsync(contextElement =>
	it.parent.box.chainAsync(listElement => f(contextElement.Url)(listElement.Url))
)

const itemIterator = it => f => it.parent.parent.box.chainAsync(contextElement =>
	it.parent.box.chainAsync(listElement =>
		it.box.chainAsync(f(contextElement.Url)(listElement.Url))
	)
)

const operateDuplicates = it => (opts = {}) => listIterator(it)(webUrl => async listUrl => {
	const items = await site(webUrl).list(listUrl).item().get({
		view: ['ID', 'FSObjType', 'Modified', ...it.box.getValues().map(prop('ID'))],
		scope: 'allItems',
		limit: MAX_ITEMS_LIMIT,
	});
	const itemsMap = new Map;
	const getHashedColumnName = itemData => {
		const values = [];
		it.box.map(column => {
			const value = itemData[column.ID];
			if (value === null || value === void 0) {
				return
			} else {
				values.push(value.get_lookupId ? value.get_lookupId() : value);
			}
		})
		return MD5(values.join('.')).toString();
	}
	for (const item of items) {
		const hashedColumnName = getHashedColumnName(item);
		if (hashedColumnName !== void 0) {
			if (!itemsMap.has(hashedColumnName)) itemsMap.set(hashedColumnName, []);
			itemsMap.get(hashedColumnName).push(item);
		}
	}
	let duplicatedsSorted = [];
	for (const [hashedColumnName, duplicateds] of itemsMap) {
		if (duplicateds.length > 1) {
			duplicatedsSorted = duplicatedsSorted.concat([duplicateds.sort((a, b) => a.Modified.getTime() - b.Modified.getTime())]);
		}
	}
	const duplicatedsFiltered = duplicatedsSorted.map(arrayInit)
	if (opts.delete) {
		await list.item([...duplicatedsFiltered.reduce((acc, el) => acc.concat(el.map(prop('ID'))), [])]).delete(opts);
	} else {
		return duplicatedsFiltered
	}
})

const report = ({ silent, actionType }) => contextBox => parentBox => box => spObjects => (
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} in ${parentBox.join()} at ${contextBox.join()}`),
	spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => async (opts = {}) => {
	let needToQuery;
	const clientContexts = {};
	const spObjectsToCache = new Map;
	const { cached, showCaml, parallelized = actionType !== 'create' } = opts;
	if (!opts.view && actionType) opts.view = ['ID'];
	const elements = await parent.parent.box.chainAsync(async contextElement => {
		let totalElements = 0;
		const contextUrl = contextElement.Url;
		const contextUrls = urlSplit(contextUrl);
		let clientContext = getClientContext(contextUrl);
		clientContexts[contextUrl] = [clientContext];
		return parent.box.chainAsync(listElement => {
			const listUrl = listElement.Url;
			let listSPObject = parent.getSPObject(listUrl)(parent.parent.getSPObject(clientContext));
			return box.chainAsync(async element => {
				if (actionType && ++totalElements >= REQUEST_BUNDLE_MAX_SIZE) {
					clientContext = getClientContext(contextUrl);
					listSPObject = parent.getSPObject(listUrl)(parent.parent.getSPObject(clientContext));
					clientContexts[contextUrl].push(clientContext);
					totalElements = 0;
				}
				listSPObject.listUrl = listUrl;
				const spObject = await spObjectGetter({
					spParentObject: actionType === 'create' ? listSPObject : getSPObject(element)(listSPObject),
					element
				});
				const isID = isNumber(element.ID);
				showCaml && spObject.camlQuery && camlLog(spObject.camlQuery.get_viewXml());
				const cachePath = isID ?
					element.ID :
					MD5(spObject.get_path().$1_1.toString().match(/<Parameters>.*<\/Parameters>/) + (opts.view || '')).toString();
				const cachePaths = [...contextUrls, 'lists', listUrl, NAME, isID ? cacheLeaf : cacheLeaf + 'Collection', cachePath];
				ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -3)(cachePaths));
				if (IS_DELETE_ACTION[actionType]) {
					needToQuery = true;
					return spObject;
				} else {
					const spObjectCached = cached ? cache.get(cachePaths) : null;
					if (cached && spObjectCached) {
						return spObjectCached;
					} else {
						needToQuery = true;
						const currentSPObjects = load(clientContext)(spObject)(opts);
						spObjectsToCache.set(cachePaths, currentSPObjects)
						return currentSPObjects;
					}
				}
			})
		})
	});
	if (needToQuery) {
		parallelized ?
			await parent.parent.box.chain(el => {
				return Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
			}) :
			await parent.parent.box.chain(async el => {
				for (let clientContext of clientContexts[el.Url]) await executorJSOM(clientContext)(opts)
			});
		!actionType && spObjectsToCache.forEach((value, key) => cache.set(value)(key))
	};
	const items = prepareResponseJSOM(opts)(elements);
	report({ ...opts, actionType })(parent.parent.box)(parent.box)(actionType === 'create' ? getInstance(Box)(items, 'item') : box)();
	return items;
}
// Inteface

export default (parent, elements) => {
	const instance = {
		box: getInstance(Box)(elements, 'item'),
		parent
	};
	const executeBinded = execute(parent)(instance.box)('properties');
	return {
		get: executeBinded(null)(prop('spParentObject')),

		create: (instance => async opts => {
			await listIterator(instance)(getColumns);
			return executeBinded('create')(async ({ spParentObject, element }) => {
				const contextUrl = spParentObject.get_context().get_url();
				const listUrl = spParentObject.listUrl;
				const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation);
				const elementNew = Object.assign({}, element);
				delete elementNew.ID;
				if (elementNew.Folder) {
					const folder = elementNew.Folder;
					delete elementNew.Folder;
					const spxFolder = site(contextUrl).list(listUrl).folder(folder);
					const folderOpts = { cached: true, silent: true };
					const folderData = await spxFolder
						.get({ ...folderOpts, asItem: true })
						.catch(async err => await spxFolder.create({ ...folderOpts, view: ['FileRef'] }));
					folderData.FileRef.split(listUrl)[1] && itemCreationInfo.set_folderUrl(folderData.FileRef);
				}
				const spObject = spParentObject.addItem(itemCreationInfo);
				return setItem(await getColumns(contextUrl)(listUrl))(elementNew)((spObject));
			})(opts)
		})(instance),

		update: (instance => async opts => {
			await listIterator(instance)(getColumns)
			return executeBinded('update')(async ({ spParentObject, element }) => {
				const contextUrl = spParentObject.get_context().get_url();
				const listUrl = spParentObject.listUrl;
				const elementNew = Object.assign({}, element);
				delete elementNew.ID;
				return setItem(await getColumns(contextUrl)(listUrl))(elementNew)(spParentObject)
			})(opts)
		})(instance),

		updateByQuery: (instance => async opts =>
			itemIterator(instance)(webUrl => listUrl => async element => {
				const { Query, Columns = {} } = element;
				if (Query === void 0) throw new Error('Query is missed');
				const list = site(webUrl).list(listUrl);
				const items = await list.item(Query).get({ ...opts, expanded: false });
				if (items.length) {
					const itemsToUpdate = items.map(item => {
						const row = Object.assign({}, Columns);
						row.ID = item.ID;
						return row
					});
					return list.item(itemsToUpdate).update(opts);
				}
			}))(instance),

		delete: (opts = {}) =>
			executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spParentObject }) =>
				overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spParentObject))(opts),

		deleteByQuery: (instance => async opts =>
			itemIterator(instance)(webUrl => listUrl => async element => {
				if (element === void 0) throw new Error('Query is missed');
				const list = site(webUrl).list(listUrl);
				const items = await list.item(element).get(opts);
				if (items.length) {
					await list.item(items.map(prop('ID'))).delete();
					return items
				}
			}))(instance),

		getDuplicates: operateDuplicates(instance),
		deleteDuplicates: (opts = {}) => operateDuplicates(instance)({ ...opts, delete: true }),

		getEmpties: (instance => (opts = {}) => {
			const columns = instance.box.getValues().map(prop('ID'));
			return listIterator(instance)(webUrl => listUrl => site(webUrl).list(listUrl).item(`${joinQueries('or')('isnull')(columns)()}`).get({
				...opts,
				scope: 'allItems',
				limit: MAX_ITEMS_LIMIT
			}))
		})(instance),

		deleteEmpties: (instance => opts =>
			listIterator(instance)(webUrl => async listUrl => site(webUrl).list(listUrl).item((await instance.getEmpties(opts)).map(prop('ID'))).delete())
		)(instance),

		merge: (instance => async (opts = {}) =>
			itemIterator(instance)(webUrl => listUrl => async element => {
				const { Key, Forced, Target, Source, Mediator = _ => identity } = element;
				const sourceColumn = Source.Column;
				const targetColumn = Target.Column;
				const list = site(webUrl).list(listUrl);
				const [sourceItems, targetItems] = await Promise.all([
					site(Source.Web || webUrl).list(Source.List || listUrl).item(concatQueries()([`${Key} IsNotNull`, Source.Query])).get({
						view: ['ID', Key, sourceColumn],
						scope: 'allItems',
						limit: MAX_ITEMS_LIMIT,
					}),
					list.item(concatQueries()([`${Key} IsNotNull`, Target.Query])).get({
						view: ['ID', Key, targetColumn],
						scope: 'allItems',
						limit: MAX_ITEMS_LIMIT,
						groupBy: Key,
					})
				])
				const itemsToMerge = [];
				await Promise.all(sourceItems.map(async  sourceItem => {
					const targetItem = targetItems[sourceItem[Key]];
					if (targetItem) {
						const itemToMerge = { ID: targetItem.ID };
						if (targetItem[targetColumn] === null || Forced) {
							itemToMerge[targetColumn] = await Mediator(sourceColumn)(sourceItem[sourceColumn]);
							itemsToMerge.push(itemToMerge);
						}
					}
				}))
				return list.item(itemsToMerge).update(opts);
			})
		)(instance),

		erase: (instance => async opts =>
			itemIterator(instance)(webUrl => listUrl => async element => {
				const { Query = '', Columns = [] } = element;
				return site(webUrl).list(listUrl).item({ Query, Columns: Columns.reduce((acc, el) => (acc[el] = null, acc), {}) }).updateByQuery(opts);
			})
		)(instance),

		createDiscussion: execute(instance.parent)(instance.box)('discussion')('create')(({ spParentObject, element }) => {
			const spObject = SP.Utilities.Utility.createNewDiscussion(spParentObject.get_context(), spParentObject, element.Title);
			for (const columnName in element) spObject.set_item(columnName, element[columnName]);
			spObject.update();
			return spObject
		}),

		createReply: execute(instance.parent)(instance.box)('reply')('create')(({ spParentObject, element }) => {
			const { Columns } = element;
			const item = getSPObject(element)(spParentObject);
			const spObject = SP.Utilities.Utility.createNewDiscussionReply(spParentObject.get_context(), item);
			for (const columnName in Columns) spObject.set_item(columnName, Columns[columnName]);
			spObject.update();
			return spObject
		})
	}
} 
