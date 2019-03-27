import {
  REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE,
  REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE,
  REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE,
  AbstractBox,
  getInstance,
  prepareResponseJSOM,
  load,
  executorJSOM,
  getInstanceEmpty,
  setItem,
  hasUrlTailSlash,
  ifThen,
  constant,
  methodEmpty,
  getSPFolderByUrl,
  pipe,
  popSlash,
  identity,
  getListRelativeUrl,
  switchCase,
  typeOf,
  isArrayFilled,
  map,
  removeEmptyUrls,
  removeDuplicatedUrls,
  shiftSlash,
  mergeSlashes,
  getParentUrl,
  deep2Iterator,
  deep3Iterator,
  listReport,
  getTitleFromUrl,
  isStrictUrl
} from './../lib/utility';
import * as cache from './../lib/cache';

import site from './../modules/site';

// Internal

const NAME = 'folder';

const getSPObject = elementUrl => ifThen(constant(elementUrl))([
  spObject => getSPFolderByUrl(elementUrl)(spObject.get_rootFolder()),
  spObject => {
    const rootFolder = methodEmpty('get_rootFolder')(spObject);
    rootFolder.isRoot = true;
    return rootFolder
  }
])

const getSPObjectCollection = elementUrl => pipe([
  getSPObject(popSlash(elementUrl)),
  methodEmpty('get_folders')
])

const liftFolderType = switchCase(typeOf)({
  object: (context) => {
    const newContext = Object.assign({}, context);
    if (!context.Url) newContext.Url = context.ServerRelativeUrl || context.FileRef;
    if (!context.ServerRelativeUrl) newContext.ServerRelativeUrl = context.Url || context.FileRef;
    return newContext
  },
  string: (contextUrl = '') => {
    const url = contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl));
    return {
      Url: url,
      ServerRelativeUrl: url,
    }
  },
  default: _ => ({
    Url: '',
    ServerRelativeUrl: '',
  })
})

class Box extends AbstractBox {
  constructor(value) {
    super(value);
    this.joinProp = 'ServerRelativeUrl';
    this.value = this.isArray
      ? ifThen(isArrayFilled)([
        pipe([
          map(liftFolderType),
          removeEmptyUrls,
          removeDuplicatedUrls
        ]),
        constant([liftFolderType()])
      ])(value)
      : liftFolderType(value);
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

// Inteface

export default parent => elements => {
  const instance = {
    box: getInstance(Box)(elements),
    parent,
  };
  return {
    get: async (opts = {}) => {
      const { asItem } = opts;
      if (asItem) opts.view = ['ListItemAllFields'];
      const { clientContexts, result } = await deep3Iterator({
        contextBox: instance.parent.parent.box,
        parentBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, parentElement, element }) => {
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
        const isCollection = hasUrlTailSlash(elementUrl);
        const spObject = isCollection
          ? getSPObjectCollection(elementUrl)(listSPObject)
          : getSPObject(elementUrl)(listSPObject);
        return load(clientContext)(spObject)(opts)
      })
      await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      return prepareResponseJSOM(opts)(result)
    },

    create: async function create(opts = {}) {
      const { asItem } = opts;
      if (asItem) opts.view = ['ListItemAllFields'];
      await cacheColumns(instance.parent.parent.box)(instance.parent.box);
      const { clientContexts, result } = await deep3Iterator({
        contextBox: instance.parent.parent.box,
        parentBox: instance.parent.box,
        elementBox: instance.box,
        bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
      })(
        ({ contextElement, clientContext, parentElement, element }) => {
          const listUrl = parentElement.Url;
          const contextUrl = contextElement.Url;
          const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element.Url);
          if (!isStrictUrl(elementUrl)) return;
          const contextSPObject = instance.parent.parent.getSPObject(clientContext);
          const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
          const itemCreationInfo = getInstanceEmpty(SP.ListItemCreationInformation);
          itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
          itemCreationInfo.set_leafName(elementUrl);
          element.Title = getTitleFromUrl(elementUrl);
          const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, element))(listSPObject.addItem(itemCreationInfo))
          return load(clientContext)(spObject.get_folder())(opts)
        })
      let needToRetry;
      let isError;
      if (instance.box.getCount()) {
        await instance.parent.parent.box.chain(async el => {
          for (const clientContext of clientContexts[el.Url]) {
            await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
              if (/This operation can only be performed on a file;/.test(err.get_message())) {
                const foldersToCreate = {};
                await deep3Iterator({
                  contextBox: instance.parent.parent.box,
                  parentBox: instance.parent.box,
                  elementBox: instance.box,
                  bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
                })(({ contextElement, parentElement, element }) => {
                  const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
                  foldersToCreate[getParentUrl(elementUrl)] = true;
                })
                await deep2Iterator({
                  contextBox: instance.parent.parent.box,
                  elementBox: instance.parent.box,
                  bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
                })(({ contextElement, element }) =>
                  site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ silent: true, expanded: true, view: ['Name'] }).then(_ => {
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
      }
      if (needToRetry) {
        return create(opts)
      } else {
        if (!isError) {
          listReport({ ...opts, NAME, actionType: 'create', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
          return prepareResponseJSOM(opts)(result);
        }
      }
    },

    update: async (opts = {}) => {
      const { asItem } = opts;
      if (asItem) opts.view = ['ListItemAllFields'];
      await cacheColumns(instance.parent.parent.box)(instance.parent.box);
      const { clientContexts, result } = await deep3Iterator({
        contextBox: instance.parent.parent.box,
        parentBox: instance.parent.box,
        elementBox: instance.box,
        bundleSize: REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE
      })(({ contextElement, clientContext, parentElement, element }) => {
        const contextUrl = contextElement.Url;
        const listUrl = parentElement.Url;
        const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element.Url);
        if (!isStrictUrl(elementUrl)) return;
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, element))(getSPObject(elementUrl)(listSPObject).get_listItemAllFields())
        return load(clientContext)(spObject.get_folder())(opts)
      })
      if (instance.box.getCount()) {
        await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      listReport({ ...opts, NAME, actionType: 'update', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
      return prepareResponseJSOM(opts)(result);
    },

    delete: async (opts = {}) => {
      const { noRecycle } = opts;
      const { clientContexts, result } = await deep3Iterator({
        contextBox: instance.parent.parent.box,
        parentBox: instance.parent.box,
        elementBox: instance.box,
        bundleSize: REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE
      })(({ contextElement, clientContext, parentElement, element }) => {
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
        if (!isStrictUrl(elementUrl)) return;
        const spObject = getSPObject(elementUrl)(listSPObject);
        !spObject.isRoot && methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
      });
      if (instance.box.getCount()) {
        await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      listReport({ ...opts, NAME, actionType: noRecycle ? 'delete' : 'recycle', box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });
      return prepareResponseJSOM(opts)(result);
    }
  }
}