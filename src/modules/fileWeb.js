import {
  AbstractBox,
  getInstance,
  methodEmpty,
  prepareResponseJSOM,
  load,
  executorJSOM,
  getContext,
  isStringEmpty,
  getParentUrl,
  prependSlash,
  convertFileContent,
  setFields,
  hasUrlTailSlash,
  mergeSlashes,
  getFolderFromUrl,
  getFilenameFromUrl,
  executorREST,
  prepareResponseREST,
  identity,
  getWebRelativeUrl,
  switchCase,
  typeOf,
  isArray,
  shiftSlash,
  ifThen,
  isArrayFilled,
  pipe,
  map,
  removeEmptyUrls,
  removeDuplicatedUrls,
  constant,
  join,
  deep2Iterator,
  deep2IteratorREST,
  webReport
} from './../lib/utility';

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

const liftFolderType = switchCase(typeOf)({
  object: (context) => {
    const newContext = Object.assign({}, context);
    if (context.Url !== '/') newContext.Url = shiftSlash(newContext.Url);
    if (!context.Url && context.ServerRelativeUrl) newContext.Url = context.ServerRelativeUrl;
    if (!context.ServerRelativeUrl && context.Url) newContext.ServerRelativeUrl = context.Url;
    return newContext
  },
  string: (contextUrl = '') => {
    const url = contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl));
    return {
      Url: url,
      ServerRelativeUrl: url
    }
  },
  default: _ => ({
    Url: '',
    ServerRelativeUrl: ''
  })
})

class Box extends AbstractBox {
  constructor(value) {
    super(value)
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

// Inteface

export default (parent, elements) => {
  const instance = {
    box: getInstance(Box)(elements),
    parent,
  };
  return {
    get: async (opts = {}) => {
      if (opts.asBlob) {
        const result = await deep2IteratorREST(instance.parent.box)(instance.box)(({ contextElement, element }) => {
          const contextUrl = contextElement.Url;
          const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
          return executorREST(contextUrl)({
            url: `${getRESTObject(elementUrl)(contextUrl)}/$value`,
            binaryStringResponseBody: true
          });
        });
        return prepareResponseREST(opts)(result);
      } else {
        if (opts.asItem) opts.view = ['ListItemAllFields'];
        const { clientContexts, result } = await deep2Iterator({
          contextBox: instance.parent.box,
          elementBox: instance.box
        })(({ contextElement, clientContext, element }) => {
          const contextUrl = contextElement.Url;
          const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
          const parentSPObject = parent.getSPObject(clientContext);
          const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl);
          const spObject = isCollection
            ? getSPObjectCollection(elementUrl)(parentSPObject)
            : getSPObject(elementUrl)(parentSPObject)
          return load(clientContext)(spObject)(opts);
        });
        await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
        return prepareResponseJSOM(opts)(result);
      }
    },

    create: async function create(opts = {}) {
      const { clientContexts, result } = await deep2Iterator({
        contextBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, element }) => {
        const {
          Content = '',
          Overwrite = true
        } = element
        if (opts.asItem) opts.view = ['ListItemAllFields'];

        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        const parentSPObject = getSPObjectCollection(getParentUrl(elementUrl))(parent.getSPObject(clientContext))
        const fileCreationInfo = new SP.FileCreationInformation;
        setFields({
          set_url: getFilenameFromUrl(elementUrl),
          set_content: convertFileContent(Content),
          set_overwrite: Overwrite
        })(fileCreationInfo)
        const spObject = parentSPObject.add(fileCreationInfo);
        return load(clientContext)(spObject)(opts);
      })
      let needToRetry;
      await instance.parent.box.chain(async el => {
        for (const clientContext of clientContexts[el.Url]) {
          await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
            if (err.get_message() === 'File Not Found.') {
              const foldersToCreate = {};
              await deep2Iterator({
                contextBox: instance.parent.box,
                elementBox: instance.box
              })(({ contextElement, element }) => {
                const elementUrl = getWebRelativeUrl(contextElement.Url)(element.Url);
                foldersToCreate[getFolderFromUrl(elementUrl)] = true;
              })
              await site(clientContext.get_url()).folder(Object.keys(foldersToCreate)).create({ silent: true, expanded: true, view: ['Name'] }).then(_ => {
                needToRetry = true;
              }).catch(identity);
            } else {
              console.error(err.get_message())
            }
          })
          if (needToRetry) break;
        }
      });
      if (needToRetry) {
        return create(opts)
      } else {
        webReport({ ...opts, NAME, actionType: 'create', box: instance.box, contextBox: instance.parent.box });
        return prepareResponseJSOM(opts)(result);
      }
    },

    update: async opts => {
      const { clientContexts, result } = await deep2Iterator({
        contextBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, element }) => {
        const { Content } = element;
        const binaryInfo = new SP.FileSaveBinaryInformation;
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        if (Content !== void 0) binaryInfo.set_content(convertFileContent(Content));
        const spObject = getSPObject(elementUrl)(parent.getSPObject(clientContext));
        spObject.saveBinary(binaryInfo);
        return spObject
      })
      await instance.parent.box.chain(async el => {
        for (const clientContext of clientContexts[el.Url]) await executorJSOM(clientContext)(opts)
      })
      webReport({ ...opts, NAME, actionType: 'update', box: instance.box, contextBox: instance.parent.box });
      return prepareResponseJSOM(opts)(result);
    },

    delete: async (opts = {}) => {
      const { noRecycle } = opts;
      const { clientContexts, result } = await deep2Iterator({
        contextBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
      });
      await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      webReport({ ...opts, NAME, actionType: 'delete', box: instance.box, contextBox: instance.parent.box });
      return prepareResponseJSOM(opts)(result);
    },

    copy: async opts => {
      const { clientContexts, result } = await deep2Iterator({
        contextBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        spObject.copyTo(getWebRelativeUrl(contextUrl)(element.To));
      });
      await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      webReport({ ...opts, NAME, actionType: 'copy', box: instance.box, contextBox: instance.parent.box });
      return prepareResponseJSOM(opts)(result);
    },

    move: async opts => {
      const { clientContexts, result } = await deep2Iterator({
        contextBox: instance.parent.box,
        elementBox: instance.box
      })(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        spObject.moveTo(getWebRelativeUrl(contextUrl)(element.To));
      });
      await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      webReport({ ...opts, NAME, actionType: 'move', box: instance.box, contextBox: instance.parent.box });
      return prepareResponseJSOM(opts)(result);
    },
  }
}
