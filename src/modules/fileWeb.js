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
  shiftSlash,
  ifThen,
  isArrayFilled,
  pipe,
  map,
  removeEmptyUrls,
  removeDuplicatedUrls,
  constant,
  deep2Iterator,
  deep2IteratorREST,
  webReport,
  removeEmptyFilenames,
  hasUrlFilename
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
  getCount() {
    return this.isArray ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0;
  }
}

const iterator = instance => deep2Iterator({
  contextBox: instance.parent.box,
  elementBox: instance.box
})

const iteratorREST = instance => deep2IteratorREST({
  contextBox: instance.parent.parent.box,
  elementBox: instance.parent.box
})

// Inteface

export default parent => elements => {
  const instance = {
    box: getInstance(Box)(elements),
    parent,
  };
  const report = actionType => (opts = {}) => webReport({ ...opts, NAME, actionType, box: instance.box, contextBox: instance.parent.box });
  return {
    get: async (opts = {}) => {
      if (opts.asBlob) {
        const result = await iteratorREST(instance)(({ contextElement, element }) => {
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
        const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
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
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        if (!hasUrlFilename(elementUrl)) return;
        const {
          Content = '',
          Overwrite = true
        } = element
        if (opts.asItem) opts.view = ['ListItemAllFields'];

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
      if (instance.box.getCount()) {
        await instance.parent.box.chain(async el => {
          for (const clientContext of clientContexts[el.Url]) {
            await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
              if (err.get_message() === 'File Not Found.') {
                const foldersToCreate = {};
                await iterator(instance)(({ contextElement, element }) => {
                  const elementUrl = getWebRelativeUrl(contextElement.Url)(element.Url);
                  foldersToCreate[getFolderFromUrl(elementUrl)] = true;
                })
                await site(clientContext.get_url()).folder(Object.keys(foldersToCreate)).create({ silentInfo: true, expanded: true, view: ['Name'] }).then(_ => {
                  needToRetry = true;
                }).catch(err => {
                  if (/already exists/.test(err.get_message())) {
                    needToRetry = true;
                  }
                });
              } else {
                if (!opts.silent && !opts.silentErrors) throw err
              }
            })
            if (needToRetry) break;
          }
        });
      }
      if (needToRetry) {
        return create(opts)
      } else {
        report('create')(opts);
        return prepareResponseJSOM(opts)(result);
      }
    },

    update: async opts => {
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
        const { Content, Url } = element;
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(Url);
        if (!hasUrlFilename(elementUrl)) return;
        const binaryInfo = new SP.FileSaveBinaryInformation;
        if (Content !== void 0) binaryInfo.set_content(convertFileContent(Content));
        const spObject = getSPObject(elementUrl)(parent.getSPObject(clientContext));
        spObject.saveBinary(binaryInfo);
        return spObject
      })
      if (instance.box.getCount()) {
        await instance.parent.box.chain(async el => {
          for (const clientContext of clientContexts[el.Url]) await executorJSOM(clientContext)(opts)
        })
      }
      report('update')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    delete: async (opts = {}) => {
      const { noRecycle } = opts;
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        if (!hasUrlFilename(elementUrl)) return;
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
      });
      if (instance.box.getCount()) {
        await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report(noRecycle ? 'delete' : 'recycle')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    copy: async opts => {
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        if (!hasUrlFilename(elementUrl)) return;
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        spObject.copyTo(getWebRelativeUrl(contextUrl)(element.To));
      });
      if (instance.box.getCount()) {
        await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report('copy')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    move: async opts => {
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, element }) => {
        const contextUrl = contextElement.Url;
        const elementUrl = getWebRelativeUrl(contextUrl)(element.Url);
        if (!hasUrlFilename(elementUrl)) return;
        const parentSPObject = instance.parent.getSPObject(clientContext);
        const spObject = getSPObject(elementUrl)(parentSPObject);
        spObject.moveTo(getWebRelativeUrl(contextUrl)(element.To));
      });
      if (instance.box.getCount()) {
        await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report('move')(opts);
      return prepareResponseJSOM(opts)(result);
    },
  }
}
