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
  getContext,
  isStringEmpty,
  prop,
  getParentUrl,
  prependSlash,
  convertFileContent,
  setFields,
  hasUrlTailSlash,
  arrayInit,
  mergeSlashes,
  getFolderFromUrl,
  getFilenameFromUrl,
  executorREST,
  prepareResponseREST,
  slice,
  identity
} from './../lib/utility';
import * as cache from './../lib/cache';

import site from './../modules/site';

//Internal

const NAME = 'file';

const getSPObject = elementUrl => spObject => {
  const folder = getFolderFromUrl(elementUrl);
  const filename = getFilenameFromUrl(elementUrl);
  const contextUrl = getContext(spObject).get_url();
  return spObject.getFileByServerRelativeUrl(mergeSlashes(`${folder ? `${contextUrl}/${folder}` : contextUrl}/${filename}`));
}

const getSPObjectCollection = elementUrl => spObject => {
  const contextUrl = getContext(spObject).get_url();
  const folder = getFolderFromUrl(elementUrl);
  return folder
    ? spObject.getFolderByServerRelativeUrl(`${contextUrl}/${folder}`).get_files()
    : spObject.get_rootFolder().get_files();
}

const getRESTObject = elementUrl => contextUrl => {
  const filename = getFilenameFromUrl(elementUrl);
  const folder = getFolderFromUrl(elementUrl);
  return mergeSlashes(`/_api/web/getfilebyserverrelativeurl('${prependSlash(folder ? `${contextUrl}/${folder}` : contextUrl)}/${filename}')`)
}


const report = ({ silent, actionType }) => parentBox => box => spObjects => (
  !silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} at ${parentBox.join()}`),
  spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => async (opts = {}) => {
  let needToQuery;
  const clientContexts = {};
  const spObjectsToCache = new Map;
  const { cached } = opts;
  if (opts.asItem) opts.view = ['ListItemAllFields'];
  const elements = await parent.box.chainAsync(async contextElement => {
    let totalElements = 0;
    const contextUrl = contextElement.Url;
    const contextUrls = urlSplit(contextUrl);
    clientContexts[contextUrl] = [getClientContext(contextUrl)];
    return box.chainAsync(async element => {
      const elementUrl = element.Url;
      let clientContext = arrayLast(clientContexts[contextUrl]);
      if (actionType && ++totalElements >= REQUEST_BUNDLE_MAX_SIZE) {
        clientContext = getClientContext(contextUrl);
        clientContexts[contextUrl].push(clientContext);
        totalElements = 0;
      }
      const parentSPObject = parent.getSPObject(clientContext);
      const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl);
      const spObject = await spObjectGetter({
        spParentObject: actionType === 'create'
          ? getSPObjectCollection(getParentUrl(elementUrl))(parentSPObject)
          : isCollection
            ? getSPObjectCollection(elementUrl)(parentSPObject)
            : getSPObject(elementUrl)(parentSPObject),
        element
      });
      const cachePath = [...contextUrls, NAME, isCollection ? cacheLeaf + 'Collection' : cacheLeaf, elementUrl];
      ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -2)(cachePath));
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

const executeREST = parent => box => cacheLeaf => actionType => restObjectGetter => async (opts = {}) => {
  const { cached, } = opts;
  const elements = await parent.box.chainAsync(async contextElement => {
    const contextUrl = contextElement.Url;
    const contextUrls = urlSplit(contextUrl);
    return box.chainAsync(async element => {
      const elementUrl = element.Url;
      const restObject = restObjectGetter({
        spParentObject: getRESTObject(elementUrl)(contextUrl),
        element
      });
      const { request, params = {} } = restObject;
      const httpProvider = params.httpProvider || executorREST(contextUrl);
      const cachePath = [...contextUrls, elementUrl || '/', NAME, cacheLeaf];
      ACTION_TYPES_TO_UNSET[actionType] && cache.unset(arrayInit(cachePath));
      const spObjectCached = cached ? cache.get(cachePath) : null;
      if (cached && spObjectCached) {
        return spObjectCached;
      } else {
        const currentSPObjects = await httpProvider(request);
        cache.set(currentSPObjects)(cachePath)
        return currentSPObjects;
      }
    })
  });
  report({ ...opts, actionType })(parent.box)(box)();
  return prepareResponseREST(opts)(elements);
}

// Inteface

export default (parent, elements) => {
  const instance = {
    box: getInstance(Box)(elements),
    parent,
  };
  const executeBinded = execute(parent)(instance.box)('properties');
  return {

    get: (instance => (opts = {}) => opts.asBlob
      ? executeREST(instance.parent)(instance.box)('properties')(null)(({ spParentObject }) => ({
        request: {
          url: `${spParentObject}/$value`,
          binaryStringResponseBody: true
        }
      }))(opts)
      : executeBinded(null)(prop('spParentObject'))(opts)
    )(instance),

    create: executeBinded('create')(async ({ spParentObject, element }) => {
      const {
        Url,
        Content = '',
        Overwrite = true
      } = element
      const folder = getFolderFromUrl(Url);
      if (folder) await site(spParentObject.get_context().get_url()).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
      const fileCreationInfo = new SP.FileCreationInformation;
      setFields({
        set_url: getFilenameFromUrl(Url),
        set_content: convertFileContent(Content),
        set_overwrite: Overwrite
      })(fileCreationInfo)
      return spParentObject.add(fileCreationInfo);
    }),

    update: executeBinded('update')(({ spParentObject, element }) => {
      const { Content } = element;
      const binaryInfo = new SP.FileSaveBinaryInformation;
      if (Content !== void 0) binaryInfo.set_content(convertFileContent(Content));
      spParentObject.saveBinary(binaryInfo);
      return spParentObject
    }),

    delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spParentObject }) =>
      overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spParentObject))(opts),

    copy: executeBinded('copy')(({ spParentObject, element }) => {
      spParentObject.copyTo(prependSlash(element.To));
      return spParentObject;
    }),

    move: executeBinded('move')(({ spParentObject, element }) => {
      spParentObject.moveTo(prependSlash(element.To));
      return spParentObject;
    })
  }
}
