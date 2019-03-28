import {
  ACTION_TYPES,
  LIBRARY_STANDART_COLUMN_NAMES,
  AbstractBox,
  getInstance,
  methodEmpty,
  prepareResponseJSOM,
  getClientContext,
  load,
  executorJSOM,
  convertFileContent,
  setFields,
  hasUrlTailSlash,
  isArray,
  mergeSlashes,
  getFolderFromUrl,
  getFilenameFromUrl,
  executorREST,
  prepareResponseREST,
  isExists,
  getSPFolderByUrl,
  popSlash,
  identity,
  getInstanceEmpty,
  executeJSOM,
  isObject,
  typeOf,
  setItem,
  ifThen,
  join,
  getListRelativeUrl,
  listReport,
  deep3Iterator,
  deep2Iterator,
  switchCase,
  shiftSlash,
  isArrayFilled,
  pipe,
  map,
  removeEmptyUrls,
  removeDuplicatedUrls,
  constant,
  deep3IteratorREST,
  deep2IteratorREST,
  stringTest,
  isUndefined,
  isBlob,
  hasUrlFilename,
  removeEmptyFilenames
} from './../lib/utility';
import axios from 'axios';
import * as cache from './../lib/cache';

import site from './../modules/site';

//Internal

const NAME = 'file';

export const getColumns = webUrl => listUrl => site(webUrl).list(listUrl).column().get({
  view: ['TypeAsString', 'InternalName', 'Title', 'Hidden'],
  groupBy: 'InternalName',
  cached: true
})

const getSPObject = elementUrl => spObject => {
  const filename = getFilenameFromUrl(elementUrl);
  const folder = getFolderFromUrl(elementUrl);
  return folder
    ? getSPFolderByUrl(folder)(spObject.get_rootFolder()).get_files().getByUrl(filename)
    : spObject.get_rootFolder().get_files().getByUrl(filename)
}

const getSPObjectCollection = elementUrl => spObject => {
  const folder = getFolderFromUrl(popSlash(elementUrl));
  return folder
    ? getSPFolderByUrl(folder)(spObject.get_rootFolder()).get_files()
    : spObject.get_rootFolder().get_files();
}

const getRESTObject = elementUrl => listUrl => contextUrl =>
  mergeSlashes(`${getRESTObjectCollection(elementUrl)(listUrl)(contextUrl)}/getbyurl('${getFilenameFromUrl(elementUrl)}')`)

const getRESTObjectCollection = elementUrl => listUrl => contextUrl => {
  const folder = getFolderFromUrl(elementUrl);
  return mergeSlashes(`/${contextUrl}/_api/web/lists/getbytitle('${listUrl}')/rootfolder${folder ? `/folders/getbyurl('${folder}')` : ''}/files`)
}

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
  getCount() {
    return this.isArray ? removeEmptyFilenames(this.value).length : hasUrlFilename(this.value[this.prop]) ? 1 : 0;
  }
}

const iterator = instance => deep3Iterator({
  contextBox: instance.parent.parent.box,
  parentBox: instance.parent.box,
  elementBox: instance.box,
  // bundleSize: REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE
})

const iteratorREST = instance => deep3IteratorREST({
  contextBox: instance.parent.parent.box,
  parentBox: instance.parent.box,
  elementBox: instance.box
})

const iteratorParent = instance => deep2Iterator({
  contextBox: instance.parent.parent.box,
  elementBox: instance.parent.box
})

const iteratorParentREST = instance => deep2IteratorREST({
  contextBox: instance.parent.parent.box,
  elementBox: instance.parent.box
})

const report = instance => actionType => opts =>
  listReport({ ...opts, NAME, actionType, box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box });

const createWithJSOM = instance => async (opts = {}) => {
  const { asItem } = opts;
  if (asItem) opts.view = ['ListItemAllFields'];
  const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, parentElement, element }) => {
    const {
      Url,
      Content = '',
      Columns = {},
      Overwrite = true
    } = element
    const listUrl = parentElement.Url;
    const contextUrl = contextElement.Url;
    const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(Url);
    if (!hasUrlFilename(elementUrl)) return;
    const contextSPObject = instance.parent.parent.getSPObject(clientContext);
    const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
    const spObjects = getSPObjectCollection(elementUrl)(listSPObject);
    const fileCreationInfo = getInstanceEmpty(SP.FileCreationInformation);
    setFields({
      set_url: `/${contextUrl}/${listUrl}/${elementUrl}`,
      set_content: convertFileContent(Content),
      set_overwrite: Overwrite
    })(fileCreationInfo)
    const fieldsToCreate = {};
    for (const fieldName in Columns) {
      const field = Columns[fieldName];
      fieldsToCreate[fieldName] = ifThen(isArray)([join(';#;#')])(field);
    }
    const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation);
    setFields({
      set_content: convertFileContent(Content),
      set_fieldValues: fieldsToCreate
    })(binaryInfo);
    const spObject = spObjects.add(fileCreationInfo);
    spObject.saveBinary(binaryInfo);
    return load(clientContext)(spObject)(opts)
  })
  let needToRetry;
  if (instance.box.getCount()) {
    await instance.parent.parent.box.chain(async el => {
      for (const clientContext of clientContexts[el.Url]) {
        await executorJSOM(clientContext)({ ...opts, silentErrors: true }).catch(async err => {
          if (err.get_message() === 'File Not Found.') {
            const foldersToCreate = {};
            await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
              const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
              foldersToCreate[getFolderFromUrl(elementUrl)] = true;
            })
            await iteratorParentREST(instance)(({ contextElement, element }) =>
              site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity)
            )
            needToRetry = true;
          } else {
            console.error(err.get_message())
          }
        })
        if (needToRetry) break;
      }
    });
  }
  if (needToRetry) {
    return createWithJSOM(instance)(opts)
  } else {
    report(instance)('create')(opts);
    return prepareResponseJSOM(opts)(result);
  }
}

const createWithREST = instance => async opts => {
  const res = await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
    const contextUrl = contextElement.Url;
    const listUrl = parentElement.Url;
    const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element.Url);
    if (!hasUrlFilename(elementUrl)) return;
    return isBlob(element.Content)
      ? createWithRESTFromBlob({ instance, contextUrl, listUrl, element })(opts)
      : createWithRESTFromString({ instance, contextUrl, listUrl, element })(opts)
  })
  report(instance)('create')(opts);
  return res;
}

const createWithRESTFromString = ({ instance, contextUrl, listUrl, element }) => async (opts = {}) => {
  const { needResponse } = opts;
  const { Url = '', Content = '', Overwrite = true, OnProgress = identity, Folder = '', Columns } = element;
  const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(Url);
  let needToRetry;
  let isError;
  const folder = getListRelativeUrl(contextUrl)(listUrl)(Folder) || getFolderFromUrl(elementUrl);
  const filename = getFilenameFromUrl(elementUrl);
  const filesUrl = getRESTObjectCollection(folder ? `${folder}/${filename}` : filename)(listUrl)(contextUrl);
  const request = {
    url: `${filesUrl}/add(url='${filename}',overwrite=${Overwrite})`,
    headers: {
      'accept': 'application/json;odata=verbose',
      'content-type': 'application/json;odata=verbose'
    },
    method: 'POST',
    data: Content,
    onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
  }
  await axios(request).catch(async err => {
    if (err.response.statusText === 'Not Found') {
      const foldersToCreate = {};
      await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
        const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
        foldersToCreate[getFolderFromUrl(elementUrl)] = true;
      })
      await iteratorParentREST(instance)(({ contextElement, element }) =>
        site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ silent: true, expanded: true, view: ['Name'] })
          .then(_ => {
            needToRetry = true;
          }).catch(identity)
      )
    } else {
      console.error(err.response.statusText)
    }
    isError = true;
  });
  if (needToRetry) {
    return createWithRESTFromString({ instance, contextUrl, listUrl, element })(opts);
  } else {
    if (!isError) {
      let response;
      if (Columns) {
        response = await site(contextUrl).library(listUrl).file({ Url: elementUrl, Columns }).update(opts)
      } else if (needResponse) {
        response = await site(contextUrl).library(listUrl).file(elementUrl).get(opts);
      }
      report(instance)('create')(opts);
      return response;
    }
  }
}

const createWithRESTFromBlob = ({ instance, contextUrl, listUrl, element }) => async opts => {
  let founds;
  const inputs = [];
  const { needResponse } = opts;
  const { Url = '', Content = '', Overwrite, OnProgress = identity, Folder = '', Columns } = element;
  const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(Url);
  let needToRetry;
  let isError;
  const folder = Folder || getFolderFromUrl(elementUrl);
  const filename = elementUrl ? getFilenameFromUrl(elementUrl) : Content.name;
  const requiredInputs = {
    __REQUESTDIGEST: true,
    __VIEWSTATE: true,
    __EVENTTARGET: true,
    __EVENTVALIDATION: true,
    ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation: true,
    ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle: true,
  }
  const listGUID = (await site(contextUrl).list(listUrl).get({ cached: true, view: 'Id' })).Id.toString();
  const res = await axios.get(`${contextUrl}/_layouts/15/Upload.aspx?List={${listGUID}}`);
  const formMatches = res.data.match(/<form(\w|\W)*<\/form>/);
  const inputRE = /<input[^<]*\/>/g;
  while (founds = inputRE.exec(formMatches)) {
    let item = founds[0];
    const id = item.match(/id=\"([^\"]+)\"/)[1];
    if (requiredInputs[id]) {
      switch (id) {
        case '__EVENTTARGET': item = item.replace(/value="[^\"]*"/, 'value="ctl00$PlaceHolderMain$ctl03$RptControls$btnOK"'); break;
        case 'ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation': item = item.replace(/value="[^\"]*"/, `value="/${folder.replace(/^\//, '')}"`); break;
        case 'ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle': if (!Overwrite) item = item.replace(/checked="[^\"]*"/, ''); break;
      }
      inputs.push(item);
    }
  }
  const form = window.document.createElement('form');
  form.innerHTML = join('')(inputs);
  const formData = new FormData(form);
  formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', Content, filename);
  const request = {
    url: `/${contextUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
    method: 'POST',
    data: formData,
    onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
  }
  const response = await axios(request);
  if (stringTest(/The selected location does not exist in this document library\./i)(response.data)) {
    const foldersToCreate = {};
    await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
      const elementUrl = getListRelativeUrl(contextElement.Url)(parentElement.Url)(element.Url);
      foldersToCreate[getFolderFromUrl(elementUrl)] = true;
    })

    isError = true;
    await iteratorParentREST(instance)(({ contextElement, element }) =>
      site(contextElement.Url).list(element.Url).folder(Object.keys(foldersToCreate)).create({ silent: true, expanded: true, view: ['Name'] })
        .then(_ => {
          needToRetry = true;
        }).catch(identity)
    )
  }
  if (needToRetry) {
    return createWithRESTFromBlob({ instance, contextUrl, listUrl, element })(opts);
  } else {
    if (!isError) {
      let response;
      if (Columns) {
        response = await site(contextUrl).library(listUrl).file({ Url: elementUrl, Columns }).update(opts)
      } else if (needResponse) {
        response = await site(contextUrl).library(listUrl).file(elementUrl).get(opts);
      }
      return response;
    }
  }
}

const copyOrMove = isMove => instance => async (opts = {}) => {
  await iteratorREST(instance)(async ({ contextElement, parentElement, element }) => {
    const { Url, To } = element;
    const contextUrl = contextElement.Url;
    const listUrl = parentElement.Url;
    const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(Url);
    if (!hasUrlFilename(elementUrl)) return;
    let targetWebUrl, targetListUrl, targetFileUrl;
    if (isObject(To)) {
      targetWebUrl = To.Web;
      targetListUrl = To.List;
      targetFileUrl = To.File || '';
    } else {
      targetWebUrl = contextUrl;
      targetListUrl = listUrl;
      targetFileUrl = To;
    }

    if (!targetWebUrl) throw new Error('Target WebUrl is missed');
    if (!targetListUrl) throw new Error('Target ListUrl is missed');
    if (!elementUrl) throw new Error('Source file Url is missed');

    const spxSourceList = site(contextUrl).list(listUrl);
    const spxSourceFile = spxSourceList.file(elementUrl);
    const spxTargetList = site(targetWebUrl).list(targetListUrl);
    const sourceFileData = await spxSourceFile.get({ asItem: true });
    const fullTargetFileUrl = /\./.test(targetFileUrl) ? targetFileUrl : (targetFileUrl + '/' + sourceFileData.FileLeafRef);
    const columnsToUpdate = {};
    for (let columnName in sourceFileData) {
      if (!LIBRARY_STANDART_COLUMN_NAMES[columnName] && sourceFileData[columnName] !== null) columnsToUpdate[columnName] = sourceFileData[columnName];
    }

    const existedColumnsToUpdate = {};
    if (Object.keys(columnsToUpdate).length) {
      for (let columnName in columnsToUpdate) {
        existedColumnsToUpdate[columnName] = sourceFileData[columnName];
      }
    }
    if (!opts.forced && contextUrl === targetWebUrl) {
      const clientContext = getClientContext(contextUrl);
      const listSPObject = instance.parent.getSPObject(listUrl)(instance.parent.parent.getSPObject(clientContext));
      const spObject = getSPObject(elementUrl)(listSPObject);
      const folder = getFolderFromUrl(targetFileUrl);
      if (folder) await site(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
      spObject[isMove ? 'moveTo' : 'copyTo'](mergeSlashes(`${targetListUrl}/${fullTargetFileUrl}`));
      await executeJSOM(clientContext)(spObject)(opts);
      await spxTargetList.file({ Url: targetFileUrl, Columns: existedColumnsToUpdate }).update({ silent: true })
    } else {
      await spxTargetList.file({
        Url: fullTargetFileUrl,
        Content: await spxSourceList.file(elementUrl).get({ asBlob: true }),
        OnProgress: element.OnProgress,
        Overwrite: element.Overwrite,
        Columns: existedColumnsToUpdate
      }).create({ silent: true });
      isMove && await spxSourceFile.delete()
    }
  })

  console.log(`${
    ACTION_TYPES[isMove ? 'move' : 'copy']} ${
    instance.parent.parent.box.getCount() * instance.parent.box.getCount() * instance.box.getCount()} ${
    NAME}(s)`);
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
      if (opts.asBlob) {
        const result = await iteratorREST(instance)(({ contextElement, parentElement, element }) => {
          const contextUrl = contextElement.Url;
          const listUrl = parentElement.Url;
          const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element.Url);
          return executorREST(contextUrl)({
            url: `${getRESTObject(elementUrl)(listUrl)(contextUrl)}/$value`,
            binaryStringResponseBody: true
          });
        });
        return prepareResponseREST(opts)(result);
      } else {
        if (opts.asItem) opts.view = ['ListItemAllFields'];
        const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, parentElement, element }) => {
          const listUrl = parentElement.Url;
          const elementUrl = getListRelativeUrl(contextElement.Url)(listUrl)(element.Url);
          const listSPObject = parent.getSPObject(listUrl)(parent.parent.getSPObject(clientContext));
          const spObject = isExists(elementUrl) && hasUrlTailSlash(elementUrl)
            ? getSPObjectCollection(elementUrl)(listSPObject)
            : getSPObject(elementUrl)(listSPObject);
          return load(clientContext)(spObject)(opts);
        });
        await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
        return prepareResponseJSOM(opts)(result)
      }
    },

    create: (opts = {}) => opts.fromString ? createWithJSOM(instance)(opts) : createWithREST(instance)(opts),

    update: async (opts = {}) => {
      const { asItem } = opts;
      if (asItem) opts.view = ['ListItemAllFields'];
      await cacheColumns(instance.parent.parent.box)(instance.parent.box);
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, parentElement, element }) => {
        const { Content, Columns, Url } = element;
        const listUrl = parentElement.Url;
        const contextUrl = contextElement.Url;
        const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(Url);
        if (!hasUrlFilename(elementUrl)) return;
        const contextSPObject = instance.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
        let spObject;
        if (isUndefined(Content)) {
          spObject = setItem(cache.get(['columns', contextUrl, listUrl]))(Object.assign({}, Columns))(getSPObject(elementUrl)(listSPObject).get_listItemAllFields())
        } else {
          const fieldsToUpdate = {};
          for (const fieldName in Columns) {
            const field = Columns[fieldName];
            fieldsToUpdate[fieldName] = ifThen(isArray)([join(';#;#')])(field);
          }
          const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation);
          setFields({
            set_content: convertFileContent(Content),
            set_fieldValues: fieldsToUpdate
          })(binaryInfo);
          spObject = getSPObject(elementUrl)(listSPObject);
          spObject.saveBinary(binaryInfo);
          spObject = spObject.get_listItemAllFields()
        }
        return load(clientContext)(spObject.get_file())(opts)
      })
      if (instance.box.getCount()) {
        await instance.parent.parent.box.chain(async el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report(instance)('update')(opts);
      return prepareResponseJSOM(opts)(result)
    },

    delete: async (opts = {}) => {
      const { noRecycle } = opts;
      const { clientContexts, result } = await iterator(instance)(({ contextElement, clientContext, parentElement, element }) => {
        const contextUrl = contextElement.Url;
        const listUrl = parentElement.Url;
        const elementUrl = getListRelativeUrl(contextUrl)(listUrl)(element.Url);
        if (!hasUrlFilename(elementUrl)) return;
        const listSPObject = instance.parent.getSPObject(listUrl)(instance.parent.parent.getSPObject(clientContext));
        const spObject = getSPObject(elementUrl)(listSPObject);
        methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
      });
      if (instance.box.getCount()) {
        await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      }
      report(instance)(noRecycle ? 'delete' : 'recycle')(opts);
      return prepareResponseJSOM(opts)(result);
    },

    copy: copyOrMove(false)(instance),

    move: copyOrMove(true)(instance),
  }
}