import {
  ACTION_TYPES_TO_UNSET,
  REQUEST_BUNDLE_MAX_SIZE,
  ACTION_TYPES,
  LIBRARY_STANDART_COLUMN_NAMES,
  Box,
  getInstance,
  methodEmpty,
  prepareResponseJSOM,
  getClientContext,
  urlSplit,
  load,
  executorJSOM,
  overstep,
  prop,
  prependSlash,
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
  slice,
  identity,
  getInstanceEmpty,
  executeJSOM,
  isObject,
  typeOf,
  setItem,
} from './../utility';
import axios from 'axios';
import * as cache from './../cache';

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
  return mergeSlashes(`${contextUrl}/_api/web/lists/getbytitle('${listUrl}')/rootfolder${folder ? `/folders/getbyurl('${folder}')` : ''}/files`)
}

const report = ({ silent, actionType }) => contextBox => parentBox => box => spObjects => (
  !silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} in ${parentBox.join()} at ${contextBox.join()}`),
  spObjects
)



const execute = parent => box => cacheLeaf => actionType => spObjectGetter => async (opts = {}) => {
  let needToQuery;
  const clientContexts = {};
  const spObjectsToCache = new Map;
  const { cached, parallelized = actionType !== 'create', asItem } = opts;
  if (asItem && !actionType) opts.view = ['ListItemAllFields'];
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
        const isCollection = isExists(elementUrl) && hasUrlTailSlash(elementUrl);
        const spParentObject = actionType === 'create'
          ? getSPObjectCollection(elementUrl)(listSPObject)
          : isCollection
            ? getSPObjectCollection(elementUrl)(listSPObject)
            : getSPObject(elementUrl)(listSPObject)
        spParentObject.listUrl = listUrl;
        const spObject = await spObjectGetter({
          spParentObject,
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
            const currentSPObjects = load(clientContext)((actionType === 'create' || actionType === 'update') && !asItem ? spObject.get_file() : spObject)(opts);
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
  report({ ...opts, actionType })(parent.parent.box)(parent.box)(box)();
  return prepareResponseJSOM(opts)(elements);
}

const executeREST = parent => box => cacheLeaf => actionType => restObjectGetter => async (opts = {}) => {
  const { cached, asItem } = opts;
  const elements = await parent.parent.box.chainAsync(contextElement => {
    const contextUrl = prependSlash(contextElement.Url);
    const contextUrls = urlSplit(contextUrl);
    return parent.box.chainAsync(listElement => {
      const listUrl = listElement.Url;
      return box.chainAsync(async element => {
        const elementUrl = element.Url;
        const fileUrl = getRESTObject(elementUrl)(listUrl)(contextUrl);
        const restObject = await restObjectGetter({
          fileUrl,
          contextUrl,
          listUrl,
          element
        });
        const { request, params = {} } = restObject;
        const httpProvider = params.httpProvider || executorREST(contextUrl);
        const isCollection = isExists(elementUrl) && hasUrlTailSlash(elementUrl);
        const cachePath = [...contextUrls, 'lists', listUrl, NAME, isCollection ? cacheLeaf + 'Collection' : cacheLeaf, elementUrl];
        ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -3)(cachePath));
        if (cached && spObjectCached) {
          return spObjectCached;
        } else {
          let currentSPObjects = await httpProvider(request);
          if (actionType === 'create' && element.Columns) {
            await axios({
              url: `${fileUrl}/ListItemAllFields`,
              headers: {
                'accept': 'application/json;odata=verbose',
                'content-type': 'application/json;odata=verbose',
                'If-Match': '*',
                'X-Http-Method': 'MERGE'
              },
              method: 'POST',
              data: JSON.stringify({
                __metadata: { type: `SP.Data.${listUrl}Item` },
                ...element.Columns
              })
            });
            currentSPObjects = await axios({
              url: `${fileUrl}${asItem ? '/ListItemAllFields' : ''}`,
              headers: {
                'accept': 'application/json;odata=verbose',
                'content-type': 'application/json;odata=verbose',
              }
            })
          }
          cache.set(currentSPObjects)(cachePath)
          return currentSPObjects;
        }
      })
    })
  });
  report({ ...opts, actionType })(parent.parent.box)(parent.box)(box)();
  return prepareResponseREST(opts)(elements);
}

const createWithJSOM = async ({ spParentObject, element }) => {
  const {
    Url,
    Content = '',
    Columns = {},
    Overwrite = true
  } = element
  const folder = getFolderFromUrl(Url);
  const listUrl = spParentObject.listUrl;
  const contextUrl = spParentObject.get_context().get_url();
  if (folder) await site(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
  const fileCreationInfo = getInstanceEmpty(SP.FileCreationInformation);
  setFields({
    set_url: `${contextUrl}/${spParentObject, listUrl}/${Url}`,
    set_content: convertFileContent(Content),
    set_overwrite: Overwrite
  })(fileCreationInfo)
  const fieldsToCreate = {};
  for (const fieldName in Columns) {
    const field = Columns[fieldName];
    fieldsToCreate[fieldName] = isArray(field) ? field.join(';#;#') : field;
  }
  const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation);
  setFields({
    set_content: convertFileContent(Content),
    set_fieldValues: fieldsToCreate
  })(binaryInfo);
  const spObject = spParentObject.add(fileCreationInfo);
  spObject.saveBinary(binaryInfo);
  return spObject.get_listItemAllFields()
}

const createWithRESTFromString = async ({ contextUrl, listUrl, element }) => {
  const { Url = '', Content = '', Overwrite = true, OnProgress = identity, Folder = '' } = element;
  const folder = Folder || getFolderFromUrl(Url);
  const filename = getFilenameFromUrl(Url);
  const filesUrl = getRESTObjectCollection(folder ? `${folder}/${filename}` : filename)(listUrl)(contextUrl);
  if (folder) await site(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
  return {
    request: {
      url: `${filesUrl}/add(url='${filename}',overwrite=${Overwrite})`,
      headers: {
        'accept': 'application/json;odata=verbose',
        'content-type': 'application/json;odata=verbose'
      },
      method: 'POST',
      data: Content,
      onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
    },
    params: {
      httpProvider: axios
    }
  }
}

const createWithRESTFromBlob = async ({ contextUrl, listUrl, element }) => {
  let founds;
  const inputs = [];
  const { Url = '', Content = '', Overwrite, OnProgress = identity, Folder = '' } = element;
  const folder = Folder || getFolderFromUrl(Url);
  const filename = Url ? getFilenameFromUrl(Url) : Content.name;
  if (folder) await site(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
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
  form.innerHTML = inputs.join('');
  const formData = new FormData(form);
  formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', Content, filename);
  return {
    request: {
      url: `${contextUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
      method: 'POST',
      data: formData,
      onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
    },
    params: {
      errorHandler: res => {
        if (/ctl00_PlaceHolderMain_LabelMessage/i.test(res.data)) throw new Error(res.data.match(/ctl00_PlaceHolderMain_LabelMessage">([^<]+)<\/span>/)[1])
      },
      httpProvider: axios
    }
  }
}

const copyOrMove = isMove => instance => async (opts = {}) => {
  const elements = await instance.parent.parent.box.chainAsync(context => {
    const contextUrl = context.Url;
    return instance.parent.box.chainAsync(list => {
      const listUrl = list.Url;
      return instance.box.chainAsync(async element => {
        const { Url, To } = element;
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
        if (!Url) throw new Error('Source file Url is missed');

        const spxSourceList = site(contextUrl).list(listUrl);
        const spxSourceFile = spxSourceList.file(Url);
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
          const spObject = getSPObject(Url)(listSPObject);
          const folder = getFolderFromUrl(targetFileUrl);
          if (folder) await site(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(identity);
          spObject[isMove ? 'moveTo' : 'copyTo'](mergeSlashes(`${targetListUrl}/${fullTargetFileUrl}`));
          await executeJSOM(clientContext)(spObject)(opts);
          await spxTargetList.file({ Url: targetFileUrl, Columns: existedColumnsToUpdate }).update({ silent: true })
        } else {
          await spxTargetList.file({
            Url: fullTargetFileUrl,
            Content: await spxSourceList.file(Url).get({ asBlob: true }),
            OnProgress: element.OnProgress,
            Overwrite: element.Overwrite,
            Columns: existedColumnsToUpdate
          }).create({ silent: true });
        }
        isMove && await spxSourceFile.delete()
      })
    })
  })
  console.log(`${
    ACTION_TYPES[isMove ? 'move' : 'copy']} ${
    instance.parent.parent.box.getLength() * instance.box.getLength()} ${
    NAME}(s)`);
  return prepareResponseJSOM(opts)(elements);
}


// Inteface

export default (parent, elements) => {
  const instance = {
    box: getInstance(Box)(elements),
    parent,
  };
  const executeBinded = execute(parent)(instance.box)('properties');
  const executeBindedREST = executeREST(parent)(instance.box)('properties');
  return {

    get: (instance => (opts = {}) => opts.asBlob
      ? executeREST(instance.parent)(instance.box)('properties')(null)(({ fileUrl }) => ({
        request: {
          url: `${fileUrl}/$value`,
          binaryStringResponseBody: true
        }
      }))(opts)
      : executeBinded(null)(prop('spParentObject'))(opts)
    )(instance),

    create: (opts = {}) => opts.fromString ?
      executeBinded('create')(createWithJSOM)(opts) :
      executeBindedREST('create')(({ contextUrl, listUrl, element }) =>
        element.Content === void 0 || typeOf(element.Content) === 'string'
          ? createWithRESTFromString({ contextUrl, listUrl, element })
          : createWithRESTFromBlob({ contextUrl, listUrl, element }))(opts),

    update: executeBinded('update')(async ({ spParentObject, element }) => {
      const { Content, Columns } = element;
      if (Content === void 0) {
        const listUrl = spParentObject.listUrl;
        const contextUrl = spParentObject.get_context().get_url();
        const elementNew = Object.assign({}, Columns);
        return setItem(await getColumns(contextUrl)(listUrl))(elementNew)((spParentObject.get_listItemAllFields()));
      } else {
        const fieldsToUpdate = {};
        for (const fieldName in Columns) {
          const field = Columns[fieldName];
          fieldsToUpdate[fieldName] = isArray(field) ? field.join(';#;#') : field;
        }
        const binaryInfo = getInstanceEmpty(SP.FileSaveBinaryInformation);
        setFields({
          set_content: Content === void 0 ? void 0 : convertFileContent(Content),
          set_fieldValues: fieldsToUpdate
        })(binaryInfo);
        spParentObject.saveBinary(binaryInfo);
        return spParentObject.get_listItemAllFields()
      }
    }),

    delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spParentObject }) =>
      overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spParentObject))(opts),

    copy: copyOrMove(false)(instance),

    move: copyOrMove(true)(instance),
  }
}