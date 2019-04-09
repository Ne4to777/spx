import {
  MAX_ITEMS_LIMIT,
  CACHE_RETRIES_LIMIT,
  AbstractBox,
  prepareResponseJSOM,
  load,
  executorJSOM,
  getInstanceEmpty,
  setItem,
  prop,
  methodEmpty,
  identity,
  deep3Iterator,
  deep2Iterator,
  listReport,
  getListRelativeUrl,
  isNull,
  isFilled,
  getInstance,
  switchCase,
  typeOf,
  shiftSlash,
  mergeSlashes,
  ifThen,
  isArrayFilled,
  map,
  constant,
  hasUrlTailSlash,
  removeEmptyNumbers,
  isNumberFilled,
  isObjectFilled,
  getArray
} from './../lib/utility';

import * as cache from './../lib/cache';
import {
  MD5
} from 'crypto-js';
import {
  getCamlView,
  craftQuery,
  camlLog,
  concatQueries
} from './../lib/query-parser';
import site from './../modules/site';

// Internal

const getPagingColumnsStr = columns => {
  if (!columns) return '';
  let str = '';
  for (let name in columns) {
    const value = columns[name];
    if (value === 'ID' || value === void 0) continue;
    str += `&p_${name}=${encodeURIComponent(value)}`
  }
  return str
}

const getTypedSPObject = typeStr => element => listUrl => parentElement => {
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
        : camlQuery.set_viewXml(getCamlView(ID));
      break;
    default:
      if (element.set_viewXml) {
        camlQuery = element;
      } else if (element.get_lookupId) {
        const spObject = parentElement.getItemById(element.get_lookupId());
        return spObject
      } else {
        const { Folder, Page } = element;
        if (Folder) {
          let fullUrl = '';
          let folderUrl = `${typeStr}${listUrl}/${Folder}`;
          if (contextUrl === '/') {
            fullUrl = folderUrl;
          } else {
            const isFullUrl = new RegExp(parentElement.get_context().get_url(), 'i').test(Folder);
            fullUrl = isFullUrl ? Folder : folderUrl;
          }
          camlQuery.set_folderServerRelativeUrl(fullUrl);
        }
        camlQuery.set_viewXml(getCamlView(element));
        if (Page) {
          const { IsPrevious, Id, Columns } = Page;
          if (Id) {
            const position = getInstanceEmpty(SP.ListItemCollectionPosition);
            position.set_pagingInfo(`${IsPrevious ? 'PagedPrev=TRUE&' : ''}Paged=TRUE&p_ID=${Id}${getPagingColumnsStr(Columns)}`);
            camlQuery.set_listItemCollectionPosition(position);
          }
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
  getCount(actionType) {
    if (this.isArray) {
      return actionType !== 'create'
        ? removeEmptyNumbers(this.value).length
        : this.value.length
    } else {
      return actionType !== 'create' ? (isNumberFilled(this.value.ID) ? 1 : 0) : 1
    }
  }
}

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
      view: ['ID', 'FSObjType', 'Modified', ...instance.box.join()]
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
      await list.item([...duplicatedsFiltered.reduce((acc, el) => acc.concat(el.map(prop('ID'))), [])]).delete({ ...opts, isSerial: true });
    } else {
      return duplicatedsFiltered
    }
  })
  return result;
}

const cacheColumns = contextBox => elementBox =>
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

const updateByQuery = instance => iterator => async (opts = {}) => {
  const module = instance.parent.NAME;
  const { result } = await iterator(async ({ contextElement, parentElement, element }) => {
    const { Folder, Query, Columns = {} } = element;
    if (Query === void 0) throw new Error('Query is missed');
    const list = site(contextElement.Url)[module](parentElement.Url);
    const items = await list.item({ Folder, Query }).get({ ...opts, expanded: false });
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
}

const deleteByQuery = instance => iterator => async opts => {
  const module = instance.parent.NAME;
  const { result } = await iterator(async ({ contextElement, parentElement, element }) => {
    if (element === void 0) throw new Error('Query is missed');
    const list = site(contextElement.Url)[module](parentElement.Url);
    const items = await list.item(element).get(opts);
    if (items.length) {
      await list.item(items.map(prop('ID'))).delete({ isSerial: true });
      return items
    }
  })
  return result;
}

const erase = instance => iterator => async opts => {
  const module = instance.parent.NAME;
  const { result } = await iterator(({ contextElement, parentElement, element }) => {
    const { Query = '', Folder, Columns } = element;
    if (!Columns) return;
    const columns = getArray(Columns);
    return site(contextElement.Url)[module](parentElement.Url).item({
      Folder,
      Query,
      Columns: columns.reduce((acc, el) => (acc[el] = null, acc), {})
    }).updateByQuery(opts);
  })
  return prepareResponseJSOM(opts)(result);
}

const getEmpties = instance => iterator => async opts => {
  const module = instance.parent.NAME;
  const columns = instance.box.value.map(prop('ID'));
  const { result } = await iterator(async ({ contextElement, element }) =>
    site(contextElement.Url)[module](element.Url).item({
      Query: `${craftQuery('or')('isnull')(columns)()}`,
      Scope: 'allItems',
      Limit: MAX_ITEMS_LIMIT,
      Folder: element.Folder
    }).get(opts))
  return result;
}

const merge = instance => iterator => async opts => {
  const module = instance.parent.NAME;
  iterator(async ({ contextElement, parentElement, element }) => {
    const { Key, Forced, To, From, Mediator = _ => identity } = element;
    if (!Key) throw new Error('Key is missed');
    const contextUrl = contextElement.Url;
    const listUrl = parentElement.Url;
    const sourceColumn = From.Column;
    if (!sourceColumn) throw new Error('From Column is missed');
    const targetColumn = To.Column;
    if (!targetColumn) throw new Error('To Column is missed');
    const list = site(contextUrl)[module](listUrl);
    const [sourceItems, targetItems] = await Promise.all([
      site(From.WebUrl || contextUrl)[module](From.List || listUrl).item({
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
}

// Inteface

export default parent => elements => {
  const instance = {
    box: getInstance(Box)(elements),
    NAME: 'item',
    parent
  }
  const getSPObject = getTypedSPObject(instance.parent.NAME === 'list' ? 'Lists/' : '');
  const iterator = deep3Iterator({
    contextBox: instance.parent.parent.box,
    parentBox: instance.parent.box,
    elementBox: instance.box
  });

  const iteratorParent = deep2Iterator({
    contextBox: instance.parent.parent.box,
    elementBox: instance.parent.box
  });

  const report = actionType => (opts = {}) =>
    listReport({ ...opts, NAME: instance.NAME, actionType, box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });

  return {
    get: async (opts = {}) => {
      const { showCaml } = opts;
      const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const spObject = getSPObject(element)(parentElement.Url)(listSPObject);
        showCaml && spObject.camlQuery && camlLog(spObject.camlQuery.get_viewXml());
        return load(clientContext)(spObject)(opts)
      })
      await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      return prepareResponseJSOM(opts)(result)
    },

    create: instance.parent.NAME === 'list'
      ? async function create(opts = {}) {
        await cacheColumns(instance.parent.parent.box)(instance.parent.box);
        const cacheUrl = ['itemCreationRetries', instance.parent.parent.box.join(), instance.parent.box.join(), instance.box.join()];
        !isNumberFilled(cache.get(cacheUrl)) && cache.set(CACHE_RETRIES_LIMIT)(cacheUrl);
        const { clientContexts, result } = await iterator(({ contextElement, clientContext, parentElement, element }) => {
          const contextUrl = contextElement.Url;
          const listUrl = parentElement.Url;
          const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation);
          const contextSPObject = instance.parent.parent.getSPObject(clientContext);
          const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
          const { Folder, Columns } = element;
          let folder = Folder;
          let newElement
          if (isObjectFilled(Columns)) {
            if (Columns.Folder) folder = Columns.Folder;
            newElement = Object.assign({}, Columns);
          } else {
            newElement = Object.assign({}, element);
          }
          delete newElement.ID;
          delete newElement.Folder;
          folder && itemCreationInfo.set_folderUrl(`/${contextUrl}/Lists/${listUrl}/${getListRelativeUrl(contextUrl)(listUrl)({ Folder: folder })}`);
          const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(newElement)(listSPObject.addItem(itemCreationInfo))
          return load(clientContext)(spObject)(opts)
        })
        let needToRetry, isError;
        await instance.parent.parent.box.chain(async el => {
          for (const clientContext of clientContexts[el.Url]) {
            await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
              if (/There is no file with URL/.test(err.get_message())) {
                const foldersToCreate = {};
                await iterator(({ contextElement, parentElement, element }) => {
                  const { Folder, Columns = {} } = element;
                  const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)({ Folder: Folder || Columns.Folder });
                  foldersToCreate[elementUrl] = true;
                })
                await iteratorParent(async ({ contextElement, element }) => {
                  const res = await site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ expanded: true, view: ['Name'] })
                    .then(_ => {
                      const cacheUrl = ['itemCreationRetries', instance.parent.parent.box.join(), instance.parent.box.join(), instance.box.join()];
                      const retries = cache.get(cacheUrl);
                      if (retries) {
                        cache.set(retries - 1)(cacheUrl)
                        return true
                      }
                    })
                    .catch(err => {
                      console.log(err);
                      if (/already exists/.test(err.get_message())) return true
                    })
                  if (res) needToRetry = true;
                })
              } else {
                throw err
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
            report('create')(opts);
            return prepareResponseJSOM(opts)(result);
          }
        }
      }
      : _ => new Promise((resolve, reject) => reject('There is no "create" method in "item" module. Use "file" module instead.')),

    update: async opts => {
      await cacheColumns(instance.parent.parent.box)(instance.parent.box);
      const { clientContexts, result } = await iterator(({ contextElement, clientContext, parentElement, element }) => {
        if (!element.Url) return;
        const contextUrl = contextElement.Url;
        const listUrl = parentElement.Url;
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const elementNew = Object.assign({}, element);
        delete elementNew.ID;
        const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(elementNew)(getSPObject(element)(parentElement.Url)(listSPObject))
        return load(clientContext)(spObject)(opts)
      })
      if (isFilled(result)) {
        await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report('update')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    updateByQuery: updateByQuery(instance)(iterator),

    delete: async (opts = {}) => {
      const { noRecycle, isSerial } = opts;
      const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
        const elementUrl = element.Url;
        if (!elementUrl) return;
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const spObject = getSPObject(element)(parentElement.Url)(listSPObject);
        !spObject.isRoot && methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject);
        return elementUrl;
      });

      if (instance.box.getCount()) {
        if (isSerial) {
          await instance.parent.parent.box.chain(async el => {
            for (const clientContext of clientContexts[el.Url]) {
              await executorJSOM(clientContext)(opts);
            }
          })
        } else {
          await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
        }
      }
      report(noRecycle ? 'delete' : 'recycle')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    deleteByQuery: deleteByQuery(instance)(iterator),

    getDuplicates: operateDuplicates(instance),
    deleteDuplicates: (opts = {}) => operateDuplicates(instance)({ ...opts, delete: true }),

    getEmpties: getEmpties(instance)(iteratorParent),
    deleteEmpties: async opts => {
      const columns = instance.box.value.map(prop('ID'));
      const { result } = await iteratorParent(async ({ contextElement, element }) => {
        const list = site(contextElement.Url).list(element.Url);
        return list.item((await list.item(columns).getEmpties(opts)).map(prop('ID'))).delete({ isSerial: true })
      })
      return result;
    },

    merge: merge(instance)(iterator),

    erase: erase(instance)(iterator),

    createDiscussion: async (opts = {}) => {
      const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
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
      report('create')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    createReply: async (opts = {}) => {
      const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
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
      report('create')(opts);
      return prepareResponseJSOM(opts)(result);
    }
  }
}