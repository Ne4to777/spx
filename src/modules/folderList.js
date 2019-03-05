import {
  ACTION_TYPES_TO_UNSET,
  REQUEST_BUNDLE_MAX_SIZE,
  ACTION_TYPES,
  Box,
  getInstance,
  prepareResponseJSOM,
  getClientContext,
  urlSplit,
  load,
  executorJSOM,
  prop,
  getInstanceEmpty,
  slice,
  setItem,
  getColumns,
  hasUrlTailSlash,
  ifThen,
  constant,
  methodEmpty,
  getSPFolderByUrl,
  pipe,
  popSlash,
  isExists,
  getParentUrl,
  identity,
  overstep
} from './../utility';
import * as cache from './../cache';

// Internal

const NAME = 'folder';

const getSPObject = elementUrl => ifThen(constant(elementUrl))([
  spObject => getSPFolderByUrl(elementUrl)(spObject.get_rootFolder()),
  methodEmpty('get_rootFolder')
])

const getSPObjectCollection = elementUrl => pipe([
  getSPObject(popSlash(elementUrl)),
  methodEmpty('get_folders')
])

const report = ({ silent, actionType }) => contextBox => parentBox => box => spObjects => (
  !silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} in ${parentBox.join()} at ${contextBox.join()}`),
  spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => async (opts = {}) => {
  let needToQuery;
  const clientContexts = {};
  const spObjectsToCache = new Map;
  const { cached, parallelized = actionType !== 'create' } = opts;
  if (opts.asItem) opts.view = ['ListItemAllFields'];
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
        const elementUrl = element.Url;
        if (actionType && ++totalElements >= REQUEST_BUNDLE_MAX_SIZE) {
          clientContext = getClientContext(contextUrl);
          listSPObject = parent.getSPObject(listUrl)(parent.parent.getSPObject(clientContext));
          clientContexts[contextUrl].push(clientContext);
          totalElements = 0;
        }
        listSPObject.listUrl = listUrl;
        listSPObject.isLibrary = opts.isLibrary;
        const isCollection = isExists(elementUrl) && hasUrlTailSlash(elementUrl);
        const spObject = await spObjectGetter({
          spParentObject: actionType === 'create'
            ? listSPObject
            : isCollection
              ? getSPObjectCollection(elementUrl)(listSPObject)
              : getSPObject(elementUrl)(listSPObject),
          element
        });
        const cachePath = [...contextUrls, 'lists', listUrl, NAME, isCollection ? cacheLeaf + 'Collection' : cacheLeaf, elementUrl];
        ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -3)(cachePath));
        if (actionType === 'delete' || actionType === 'recycle') {
          needToQuery = true;
        } else {
          const spObjectCached = cached ? cache.get(cachePath) : null;
          if (cached && spObjectCached) {
            return spObjectCached;
          } else {
            needToQuery = true;
            const currentSPObjects = load(clientContext)(spObject)(opts);
            spObjectsToCache.set(cachePath, currentSPObjects)
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
    box: getInstance(Box)(elements, 'folder'),
    parent,
  };
  const executeBinded = execute(parent)(instance.box)('properties');
  return {
    get: executeBinded(null)(prop('spParentObject')),

    create: executeBinded('create')(async ({ spParentObject, element }) => {
      const listUrl = spParentObject.listUrl;
      const contextUrl = spParentObject.get_context().get_url();
      const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation);
      const parentFolderUrl = getParentUrl(element.Url);
      const elementNew = Object.assign({}, element);
      parentFolderUrl && await spx(contextUrl).list(listUrl).folder(parentFolderUrl).create({
        view: ['FileLeafRef'],
        silent: true,
        expanded: true,
      }).catch(identity);
      itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
      itemCreationInfo.set_leafName(element.Url);
      const spObject = spParentObject.addItem(itemCreationInfo);
      return setItem(await getColumns(contextUrl)(listUrl))(elementNew)((spObject));
    }),

    delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spParentObject }) =>
      overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spParentObject))(opts)
  }
}