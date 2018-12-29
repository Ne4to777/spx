import * as cache from './../cache';
import spx from './../modules/site';

export const REQUEST_TIMEOUT = 3600000;
export const MAX_ITEMS_LIMIT = 100000;
export const REQUEST_BUNDLE_MAX_SIZE = 252;
export const ACTION_TYPES = {
  create: 'created',
  update: 'updated',
  delete: 'deleted',
  recycle: 'recycled',
  get: 'get',
  copy: 'copied',
  move: 'moved',
  restore: 'restored',
  clear: 'cleared'
}

export const CHUNK_MAX_LENGTH = 500;
export const TYPES = {
  list: [{
    name: 'genericList',
    description: 'Custom list'
  }, {
    name: 'documentLibrary',
    description: 'Document library'
  }, {
    name: 'pictureLibrary',
    description: 'Picture library'
  }],
  column: ['Text', 'Note', 'Number', 'Choice', 'DateTime', 'Lookup', 'User'],
  listCodes: {
    100: 'genericList',
    101: 'documentLibrary'
  }
}

export const ACTION_TYPES_TO_UNSET = {
  create: true,
  update: true,
  delete: true,
  recycle: true,
  restore: true
}

export const FILE_LIST_TEMPLATES = {
  101: true,
  109: true,
  110: true,
  113: true,
  114: true,
  116: true,
  119: true,
  121: true,
  122: true,
  123: true,
  175: true,
  851: true,
  10102: true
}

export const LIBRARY_STANDART_COLUMN_NAMES = {
  AppAuthor: true,
  AppEditor: true,
  Author: true,
  CheckedOutTitle: true,
  CheckedOutUserId: true,
  CheckoutUser: true,
  ContentTypeId: true,
  Created: true,
  Created_x0020_By: true,
  Created_x0020_Date: true,
  DocConcurrencyNumber: true,
  Editor: true,
  FSObjType: true,
  FileDirRef: true,
  FileLeafRef: true,
  FileRef: true,
  File_x0020_Size: true,
  File_x0020_Type: true,
  FolderChildCount: true,
  GUID: true,
  HTML_x0020_File_x0020_Type: true,
  ID: true,
  InstanceID: true,
  IsCheckedoutToLocal: true,
  ItemChildCount: true,
  Last_x0020_Modified: true,
  MetaInfo: true,
  Modified: true,
  Modified_x0020_By: true,
  Order: true,
  ParentLeafName: true,
  ParentVersionString: true,
  ProgId: true,
  ScopeId: true,
  SortBehavior: true,
  SyncClientId: true,
  TemplateUrl: true,
  Title: true,
  UniqueId: true,
  VirusStatus: true,
  WorkflowInstanceID: true,
  WorkflowVersion: true,
  owshiddenversion: true,
  xd_ProgID: true,
  xd_Signature: true,
  _CheckinComment: true,
  _CopySource: true,
  _HasCopyDestinations: true,
  _IsCurrentVersion: true,
  _Level: true,
  _ModerationComments: true,
  _ModerationStatus: true,
  _SharedFileIndex: true,
  _SourceUrl: true,
  _UIVersion: true,
  _UIVersionString: true,
}

export const prepareResponseJSOM = (inputData, opts = {}) => {
  if (!inputData) return inputData;
  const { expanded, groupBy } = opts;
  const getValues = spObject => {
    if (!spObject) return;
    if (opts.asItem) spObject = spObject.get_listItemAllFields();
    if (spObject.get_fieldValues) {
      return spObject.get_fieldValues();
    } else if (spObject.get_objectData) {
      return spObject.get_objectData().get_properties();
    } else {
      throw new Error(`Wrong spObject object: ${spObject ? JSON.stringify(spObject) : spObject}`);
    }
  }
  if (typeOf(inputData) === 'array') {
    const flattenElements = flatter(inputData);
    const isArray = opts.isArray || flattenElements.length > 1;
    if (expanded) return isArray ? flattenElements : flattenElements[0];
    const preparedElements = flattenElements.map(getValues)
    return groupBy ? groupper(groupBy, preparedElements) : isArray ? preparedElements : preparedElements[0];
  } else {
    return getValues(inputData)
  }
}

export const prepareResponseREST = (inputData, opts = {}) => {
  if (!inputData) return inputData;
  const { expanded, groupBy } = opts;
  const getValues = restObject => {
    if (!restObject) return;
    try {
      const result = JSON.parse(restObject.body).d;
      return result.results ? result.results : result;
    } catch (err) {
      return restObject.body
    }
  }
  if (typeOf(inputData) === 'array') {
    const flattenElements = flatter(inputData);
    const isArray = opts.isArray || flattenElements.length > 1;
    if (expanded) return isArray ? flattenElements : flattenElements[0];
    const preparedElements = flattenElements.map(getValues);
    return groupBy ? groupper(groupBy, preparedElements) : isArray ? preparedElements : preparedElements[0];
  } else {
    return getValues(inputData)
  }
}

export const getClientContext = contextUrl => {
  const clientContext = new SP.ClientContext(contextUrl);
  clientContext.set_requestTimeout(REQUEST_TIMEOUT);
  return clientContext;
}

export const load = (clientContext, data, opts = {}) => {
  let { groupBy, view } = opts;
  if (view) {
    if (typeOf(view) === 'array') {
      if (!view.length) view = void 0;
    } else {
      view = [view];
    }
    if (groupBy) view = view.concat(typeOf(groupBy) === 'array' ? groupBy : [groupBy]);
  };
  if (data.getEnumerator) {
    return clientContext.loadQuery.apply(clientContext, view ? [data, `Include(${view})`] : [data])
  } else {
    clientContext.load.apply(clientContext, view ? [data, view] : [data])
    return data;
  }
}

export const executeQueryAsync = (clientContext, opts = {}) =>
  new Promise((resolve, reject) => {
    clientContext.executeQueryAsync((sender, args) => {
      resolve([sender, args])
    }, (sender, args) => {
      if (!opts.silent && !opts.silentErrors) {
        console.error(`\nMessage: ${
          args.get_message().replace(/\n{1,}/g, ' ')}\nValue: ${
          args.get_errorValue()}\nType: ${
          args.get_errorTypeName()}\nId: ${
          args.get_errorTraceCorrelationId()}`);
      }
      reject(args)
    })
  })

export const requestExecutor = (contextUrl = '/', params = {}) =>
  new Promise((resolve, reject) => {
    new SP.RequestExecutor(contextUrl).executeAsync({
      ...params,
      method: !params.method || /^get$/i.test(params.method) ? 'GET' : 'POST',
      success: resolve,
      error: reject
    })
  })

export const showCache = () => cache().show();

export const isGUID = value => /^[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}$/.test(value);

export const getValidMethodName = (spObject, methodName) => {
  let firstLetter = methodName.substring(0, 1);
  let restLetters = methodName.substring(1);
  let methodName1 = `get_${firstLetter.toLowerCase()}${restLetters}`;
  let methodName2 = `get${firstLetter.toUpperCase()}${restLetters}`;
  return spObject[methodName1] ? methodName1 : methodName2;
}

export const convertFileContent = data => {
  if (typeOf(data) === 'arraybuffer') {
    let out = '';
    const bytes = new Uint8Array(data);
    for (let byte of bytes) out += String.fromCharCode(byte);
    return btoa(out);
  } else {
    try {
      return btoa(data);
    } catch (err) {
      return data;
    }
  }
}

export const flatter = base => base.reduce((acc, el) => acc.concat(typeOf(el) === 'array' ? flatter(el) : el), [])

export const groupper = (by, array) => {
  const groupBys = typeOf(by) === 'array' ? by : [by];
  const groupFlat = by => array => {
    const groupped = {};
    array.map(el => {
      const elValue = el[by];
      const groupValue = groupped[elValue];
      groupped[elValue] = groupValue ? (typeOf(groupValue) === 'array' ? groupValue.concat(el) : [groupValue, el]) : el;
    })
    return groupped
  }

  const mapper = (array, fn) => {
    const result = {};
    const mapperR = (acc, el, prop) => {
      if (prop) {
        const childEl = el[prop];
        if (typeOf(childEl) === 'array') {
          acc[prop] = fn(childEl);
          return acc;
        } else {
          acc[prop] = childEl;
          if (typeOf(childEl) === 'object') {
            for (let childProp in childEl) {
              acc[prop] = childEl;
              mapperR(acc[prop], childEl, childProp);
            }
          }
          return acc;
        }
      } else {
        if (typeOf(el) === 'array') {
          return fn(el)
        } else {
          for (let prop in el) mapperR(acc, el, prop);
          return acc;
        }
      }
    }
    return mapperR(result, array);
  }

  const merger = (obj, fn) => groupBys.reduce((acc, el) => mapper(acc, fn(el)), Object.assign(obj));
  return merger(array, groupFlat);
}

export const setItem = (fieldsInfo, item, fields) => {
  for (let fieldsTitle in fields) {
    const fieldValues = fields[fieldsTitle];
    if (fieldValues === void 0) continue;
    const fieldInfo = fieldsInfo[fieldsTitle];
    if (fieldInfo) {
      if (fieldInfo.ReadOnlyField) continue;
      const setItem = value => item.set_item(fieldInfo.InternalName, value);
      switch (fieldInfo.TypeAsString) {
        case 'Text':
        case 'Note':
        case 'Number':
        case 'Boolean':
        case 'Choice':
        case 'URL':
        case 'ModStat':
        case 'File':
        case 'DateTime':
          setItem(fieldValues);
          break;
        case 'Lookup':
          setItem(_setLookup(new SP.FieldLookupValue, fieldValues));
          break;
        case 'LookupMulti':
          setItem(_setLookupMulti(new SP.FieldLookupValue, fieldValues));
          break;
        case 'User':
          setItem(_setLookup(new SP.FieldUserValue, fieldValues));
          break;
        case 'UserMulti':
          setItem(_setLookupMulti(new SP.FieldUserValue, fieldValues));
          break;
      }
    } else {
      throw new Error(`Missed fields: ${fieldsTitle}`);
    }
  }
  item.update();
}

export const _setLookupMulti = (valueType, fieldValues) => {
  const values = fieldValues.reduce((acc, value) => (value && acc.push(this._setLookup(valueType, value)), acc), []);
  return values.length ? values : null;
}

export const _setLookup = (lookupValue, fieldValues) => {
  if (fieldValues !== null) {
    if (fieldValues.get_lookupId) {
      lookupValue.set_lookupId(fieldValues.get_lookupId());
    } else {
      const valueInt = parseInt(fieldValues);
      if (typeOf(valueInt) === 'number' && valueInt > 0) {
        lookupValue.set_lookupId(valueInt);
      } else {
        lookupValue = null;
      }
    }
  }
  return lookupValue;
}

export const initFieldsInfo = async (contextUrls, listUrls, targetObject) => {
  return Promise.all(contextUrls.reduce((acc, contextUrl) => {
    if (!targetObject[contextUrl]) targetObject[contextUrl] = {};
    return acc.concat(listUrls.map(async listUrl => {
      if (!targetObject[contextUrl][listUrl]) {
        targetObject[contextUrl][listUrl] = await spx(contextUrl).list(listUrl).column().get({
          view: ['TypeAsString', 'InternalName', 'Title', 'ReadOnlyField'],
          groupBy: 'InternalName',
          cached: true
        })
      }
    }))
  }, []))
}

export const setField = (spObject, method, value) => value !== void 0 && spObject[method](value);

export const setFields = (spObject, params) => {
  for (let method in params) setField(spObject, method, params[method]);
}

export const getLastPath = url => url.split('/').slice(-1)[0];

export const getSPFolderByUrl = (spFolderObject, url) => {
  const folderUrls = url.split('/').filter(el => !!el.length);
  const getSPFolderR = (base, i) =>
    i < folderUrls.length ? (getSPFolderR(base.get_folders().getByUrl(folderUrls[i]), ++i)) : base;
  return getSPFolderR(spFolderObject, 0);
}

export const getFolderAndFilenameFromUrl = elementUrl => {
  let folder, filename;
  if (elementUrl) {
    const fileSplits = elementUrl.split('/');
    if (/\./.test(fileSplits.slice(-1))) filename = fileSplits.pop();
    folder = fileSplits.join('/');
  }
  return { filename, folder }
}