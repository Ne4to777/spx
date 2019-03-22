import {
	MAX_ITEMS_LIMIT,
	AbstractBox,
	getInstance,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstanceEmpty,
	setItem,
	prop,
	methodEmpty,
	identity,
	arrayInit,
	typeOf,
	switchCase,
	deep3Iterator,
	hasUrlTailSlash,
	shiftSlash,
	mergeSlashes,
	ifThen,
	isArrayFilled,
	map,
	constant,
	deep2Iterator,
	listReport,
	getListRelativeUrl,
	popSlash,
	isDefined,
	isNull
} from './../lib/utility';
import * as cache from './../lib/cache';
import {
	MD5
} from 'crypto-js';
import {
	getCamlView,
	joinQueries,
	camlLog,
	concatQueries
} from './../lib/query-parser';
import site from './../modules/site';

// Internal

const NAME = 'item';

const getSPObject = element => listUrl => parentElement => {
	let camlQuery = new SP.CamlQuery();
	const ID = element.ID;
	const contextUrl = parentElement.get_context().get_url();
	switch (typeOf(ID)) {
		case 'number': {
			const spObject = parentElement.getItemById(ID);
			return spObject
		}
		case 'string':
			hasUrlTailSlash(ID)
				? camlQuery.set_folderServerRelativeUrl(`${contextUrl}/${popSlash(ID)}`)
				: camlQuery.set_viewXml(getCamlView()(ID));
			break;
		default:
			if (element.set_viewXml) {
				camlQuery = element;
			} else if (element.get_lookupId) {
				const spObject = parentElement.getItemById(element.get_lookupId());
				return spObject
			} else {
				const { Folder, Query, Page, IsLibrary } = element;
				if (Folder) {
					let fullUrl = '';
					if (contextUrl === '/') {
						fullUrl = (IsLibrary ? '' : 'Lists/') + `${listUrl}/${Folder}`;
					} else {
						const isFullUrl = new RegExp(parentElement.get_context().get_url(), 'i').test(Folder);
						fullUrl = isFullUrl ? Folder : ((IsLibrary ? '' : 'Lists/') + `${listUrl}/${Folder}`);
					}
					camlQuery.set_folderServerRelativeUrl(fullUrl);
				}
				camlQuery.set_viewXml(getCamlView(element)(Query));
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
	return spObject
}

const liftItemType = switchCase(typeOf)({
	object: item => {
		const newItem = Object.assign({}, item);
		if (!item.Url) newItem.Url = item.ID;
		if (!item.ID) newItem.ID = item.Url;
		return newItem
	},
	string: (item = '') => {
		const url = shiftSlash(mergeSlashes(item));
		return {
			Url: url,
			ID: url,
		}
	},
	number: item => ({
		ID: item,
		Url: item
	}),
	default: _ => ({
		ID: void 0,
		Url: void 0
	})
})

class Box extends AbstractBox {
	constructor(value) {
		super(value);
		this.joinProp = 'ID';
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				map(liftItemType),
				constant([liftItemType()])
			])(value)
			: liftItemType(value);
	}
}

export const cacheColumns = contextBox => elementBox =>
	deep2Iterator({ contextBox, elementBox })(async ({ contextElement, element }) => {
		const contextUrl = contextElement.Url;
		const listUrl = element.Url;
		if (!cache.get(['columns', contextUrl, listUrl])) {
			const columns = await site(contextUrl).list(listUrl).column().get({
				view: ['TypeAsString', 'InternalName', 'Title', 'Sealed'],
				groupBy: 'InternalName'
			})
			cache.set(columns)(['columns', contextUrl, listUrl]);
		}
	})

const operateDuplicates = instance => async (opts = {}) => {
	const { result } = await deep2Iterator({
		contextBox: instance.parent.parent.box,
		elementBox: instance.parent.box
	})(async ({ contextElement, element }) => {
		const list = site(contextElement.Url).list(element.Url);
		const items = await list.item({
			Scope: 'allItems',
			Limit: MAX_ITEMS_LIMIT
		}).get({
			view: ['ID', 'FSObjType', 'Modified', ...instance.box.value.map(prop('ID'))]
		});
		const itemsMap = new Map;
		const getHashedColumnName = itemData => {
			const values = [];
			instance.box.chain(column => {
				const value = itemData[column.ID];
				isDefined(value) && values.push(value.get_lookupId ? value.get_lookupId() : value);
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
	return result;
}

// Inteface

export default (parent, elements) => {
	const instance = {
		box: getInstance(Box)(elements),
		parent
	};
	return {
		get: async (opts = {}) => {
			const { showCaml } = opts;
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(({ clientContext, parentElement, element }) => {
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const spObject = getSPObject(element)(parentElement.Url)(listSPObject);
				showCaml && spObject.camlQuery && camlLog(spObject.camlQuery.get_viewXml());
				return load(clientContext)(spObject)(opts)
			})
			await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			return prepareResponseJSOM(opts)(result)
		},

		create: async function create(opts = {}) {
			await cacheColumns(instance.parent.parent.box)(instance.parent.box);
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(
				({ contextElement, clientContext, parentElement, element }) => {
					const contextUrl = contextElement.Url;
					const listUrl = parentElement.Url;
					const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation);
					const contextSPObject = instance.parent.parent.getSPObject(clientContext);
					const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
					const { Folder, Columns } = element;
					const newElement = Object.assign({}, Columns || element);
					delete newElement.ID;
					delete newElement.Folder;
					Folder && itemCreationInfo.set_folderUrl(`/${contextUrl}${opts.IsLibrary ? '/' : '/Lists/'}${listUrl}/${getListRelativeUrl(contextUrl)(listUrl)(Folder)}`);
					const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(newElement)(listSPObject.addItem(itemCreationInfo))
					return load(clientContext)(spObject)(opts)
				})
			let needToRetry;
			let isError;
			await instance.parent.parent.box.chain(async el => {
				for (const clientContext of clientContexts[el.Url]) {
					await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
						if (/There is no file with URL/.test(err.get_message())) {
							const foldersToCreate = {};
							await deep3Iterator({
								contextBox: instance.parent.parent.box,
								parentBox: instance.parent.box,
								elementBox: instance.box
							})(({ contextElement, parentElement, element }) => {
								const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Folder);
								foldersToCreate[elementUrl] = true;
							})
							await deep2Iterator({
								contextBox: instance.parent.parent.box,
								elementBox: instance.parent.box
							})(({ contextElement, element }) =>
								site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ expanded: true, view: ['Name'] }).then(_ => {
									needToRetry = true;
								}).catch(identity)
							)
						} else {
							!opts.silent && !opts.silentErrors && console.error(err.get_message())
						}
						isError = true;
					})
					if (needToRetry) break;
				}
			});
			if (needToRetry) {
				return create(opts)
			} else {
				if (!isError) {
					listReport({ ...opts, NAME, actionType: 'create', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
					return prepareResponseJSOM(opts)(result);
				}
			}
		},

		update: async opts => {
			await cacheColumns(instance.parent.parent.box)(instance.parent.box);
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(({ contextElement, clientContext, parentElement, element }) => {
				if (!element.ID) throw new Error('update failed. ID is missed');
				const contextUrl = contextElement.Url;
				const listUrl = parentElement.Url;
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const elementNew = Object.assign({}, element);
				delete elementNew.ID;
				const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(elementNew)(getSPObject(element)(parentElement.Url)(listSPObject))
				return load(clientContext)(spObject)(opts)
			})
			await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			listReport({ ...opts, NAME, actionType: 'update', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
			return prepareResponseJSOM(opts)(result);
		},

		updateByQuery: async opts => {
			const { result } = deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(async ({ contextElement, parentElement, element }) => {
				const { Query, Columns = {} } = element;
				if (Query === void 0) throw new Error('Query is missed');
				const list = site(contextElement.Url).list(parentElement.Url);
				const items = await list.item(Query).get({ ...opts, expanded: false });
				if (items.length) {
					const itemsToUpdate = items.map(item => {
						const row = Object.assign({}, Columns);
						row.ID = item.ID;
						return row
					});
					return list.item(itemsToUpdate).update(opts);
				}
			})
			return result;
		},

		delete: async (opts = {}) => {
			const { noRecycle } = opts;
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box,
			})(({ clientContext, parentElement, element }) => {
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const spObject = getSPObject(element)(parentElement.Url)(listSPObject);
				!spObject.isRoot && methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			});
			await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			listReport({ ...opts, NAME, actionType: 'delete', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
			return prepareResponseJSOM(opts)(result);
		},

		deleteByQuery: async opts => {
			const { result } = deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(async ({ contextElement, parentElement, element }) => {
				if (element === void 0) throw new Error('Query is missed');
				const list = site(contextElement.Url).list(parentElement.Url);
				const items = await list.item(element).get(opts);
				if (items.length) {
					await list.item(items.map(prop('ID'))).delete();
					return items
				}
			})
			return result;
		},

		getDuplicates: operateDuplicates(instance),
		deleteDuplicates: (opts = {}) => operateDuplicates(instance)({ ...opts, delete: true }),

		getEmpties: async opts => {
			const columns = instance.box.value.map(prop('ID'));
			const { result } = await deep2Iterator({
				contextBox: instance.parent.parent.box,
				elementBox: instance.parent.box
			})(async ({ contextElement, element }) =>
				site(contextElement.Url).list(element.Url).item({
					Query: `${joinQueries('or')('isnull')(columns)()}`,
					Scope: 'allItems',
					Limit: MAX_ITEMS_LIMIT
				}).get(opts))
			return result;
		},
		deleteEmpties: async opts => {
			const columns = instance.box.value.map(prop('ID'));
			const { result } = await deep2Iterator({
				contextBox: instance.parent.parent.box,
				elementBox: instance.parent.box
			})(async ({ contextElement, element }) => {
				const list = site(contextElement.Url).list(element.Url);
				return list.item((await list.item(columns).getEmpties(opts)).map(prop('ID'))).delete()
			})
			return result;
		},

		merge: async (opts = {}) => {
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(async ({ contextElement, clientContext, parentElement, element }) => {
				const { Key, Forced, To, From, Mediator = _ => identity } = element;
				if (!Key) throw new Error('Key is missed');
				const contextUrl = contextElement.Url;
				const listUrl = parentElement.Url;
				const sourceColumn = From.Column;
				if (!sourceColumn) throw new Error('From Column is missed');
				const targetColumn = To.Column;
				if (!targetColumn) throw new Error('To Column is missed');
				const list = site(contextUrl).list(listUrl);
				const [sourceItems, targetItems] = await Promise.all([
					site(From.Web || contextUrl).list(From.List || listUrl).item({
						Query: concatQueries()([`${Key} IsNotNull`, From.Query]),
						Scope: 'allItems',
						Limit: MAX_ITEMS_LIMIT
					}).get({
						view: ['ID', Key, sourceColumn]
					}),
					list.item({
						Query: concatQueries()([`${Key} IsNotNull`, To.Query]),
						Scope: 'allItems',
						Limit: MAX_ITEMS_LIMIT,
					}).get({
						view: ['ID', Key, targetColumn],
						groupBy: Key
					})
				])
				const itemsToMerge = [];
				await Promise.all(sourceItems.map(async  sourceItem => {
					const targetItemGroup = targetItems[sourceItem[Key]];
					if (targetItemGroup) {
						const targetItem = targetItemGroup[0];
						const itemToMerge = { ID: targetItem.ID };
						if (isNull(targetItem[targetColumn]) || Forced) {
							itemToMerge[targetColumn] = await Mediator(sourceColumn)(sourceItem[sourceColumn]);
							itemsToMerge.push(itemToMerge);
						}
					}
				}))
				return itemsToMerge.length && list.item(itemsToMerge).update(opts);
			})
		},

		erase: async opts => {
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(({ contextElement, parentElement, element }) => {
				const { Query = '', Columns = ['Title'] } = element;
				return site(contextElement.Url).list(parentElement.Url).item({ Query, Columns: Columns.reduce((acc, el) => (acc[el] = null, acc), {}) }).updateByQuery(opts);
			})
			await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			listReport({ ...opts, NAME, actionType: 'erase', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
			return prepareResponseJSOM(opts)(result);
		},

		createDiscussion: async (opts = {}) => {
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(({ clientContext, parentElement, element }) => {
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				delete element.Url;
				delete element.ID;
				const spObject = SP.Utilities.Utility.createNewDiscussion(clientContext, listSPObject, element.Title);
				for (const columnName in element) spObject.set_item(columnName, element[columnName]);
				spObject.update();
				return spObject
			})
			await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			listReport({ ...opts, NAME, actionType: 'create', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
			return prepareResponseJSOM(opts)(result);
		},

		createReply: async (opts = {}) => {
			const { clientContexts, result } = await deep3Iterator({
				contextBox: instance.parent.parent.box,
				parentBox: instance.parent.box,
				elementBox: instance.box
			})(({ clientContext, parentElement, element }) => {
				const { Columns } = element;
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const spItemObject = getSPObject(element)(parentElement.Url)(listSPObject);
				const spObject = SP.Utilities.Utility.createNewDiscussionReply(clientContext, spItemObject);
				for (const columnName in Columns) spObject.set_item(columnName, Columns[columnName]);
				spObject.update();
				return spObject
			})
			await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			listReport({ ...opts, NAME, actionType: 'create', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
			return prepareResponseJSOM(opts)(result);
		},
	}
} 
