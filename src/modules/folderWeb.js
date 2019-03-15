import {
  ACTION_TYPES_TO_UNSET,
  REQUEST_BUNDLE_MAX_SIZE,
  ACTION_TYPES,
  Box,
  getInstance,
  methodEmpty,
  prepareResponseJSOM,
  getClientContext,
  urlSplit,
  load,
  executorJSOM,
  overstep,
  getTitleFromUrl,
  identity,
  getContext,
  isStringEmpty,
  popSlash,
  prop,
  getParentUrl,
  method,
  ifThen,
  constant,
  pipe,
  hasUrlTailSlash,
  slice
} from './../utility';
import * as cache from './../cache';
import site from './../modules/site';

//Internal

const NAME = 'folder';

const getSPObject = elementUrl => ifThen(constant(elementUrl))([
  method('getFolderByServerRelativeUrl')(elementUrl),
  methodEmpty('get_rootFolder')
])

const getSPObjectCollection = elementUrl => pipe([
  ifThen(constant(!elementUrl || elementUrl === '/'))([
    getSPObject(),
    getSPObject(popSlash(elementUrl))
  ]),
  methodEmpty('get_folders')
])

const report = ({ silent, actionType }) => parentBox => box => spObjects => (
  !silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} at ${parentBox.join()}`),
  spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => async (opts = {}) => {
  let needToQuery;
  const clientContexts = {};
  const spObjectsToCache = new Map;
  const { cached, } = opts;
  const elements = await parent.box.chainAsync(async contextElement => {
    let totalElements = 0;
    const contextUrl = contextElement.Url;
    const contextUrls = urlSplit(contextUrl);
    let clientContext = getClientContext(contextUrl);
    let parentSPObject = parent.getSPObject(clientContext);
    clientContexts[contextUrl] = [clientContext];
    return box.chainAsync(async element => {
      const elementUrl = element.Url;
      if (actionType && ++totalElements >= REQUEST_BUNDLE_MAX_SIZE) {
        clientContext = getClientContext(contextUrl);
        parentSPObject = parent.getSPObject(clientContext);
        clientContexts[contextUrl].push(clientContext);
        totalElements = 0;
      }
      const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl);
      const spObject = await spObjectGetter({
        spParentObject: actionType === 'create'
          ? getSPObject(getParentUrl(elementUrl))(parentSPObject)
          : isCollection
            ? getSPObjectCollection(elementUrl)(parentSPObject)
            : getSPObject(elementUrl)(parentSPObject),
        elementUrl
      });
      const cachePath = [...contextUrls, NAME, isCollection ? cacheLeaf + 'Collection' : cacheLeaf, elementUrl];
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
  });
  if (needToQuery) {
    await parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
    !actionType && spObjectsToCache.forEach((value, key) => cache.set(value)(key))
  };
  report({ ...opts, actionType })(parent.box)(box)();
  return prepareResponseJSOM(opts)(elements);
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

    create: (instance => executeBinded('create')(async ({ spParentObject, elementUrl }) => {
      const clientContext = getContext(spParentObject);
      const parentFolderUrl = getParentUrl(elementUrl);
      if (parentFolderUrl) await site(clientContext.get_url()).folder(parentFolderUrl).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
      return getSPObjectCollection(`${parentFolderUrl}/`)(instance.parent.getSPObject(clientContext)).add(getTitleFromUrl(elementUrl));
    }))(instance),

    delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spParentObject }) =>
      overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spParentObject))(opts)
  }
}