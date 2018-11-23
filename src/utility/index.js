import * as cache from './../cache';
import axios from 'axios';
import PathBuilder from './../pathBuilder';
import spx from './../modules/site';

export const REQUEST_TIMEOUT = 3600000;

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

export const prepareResponseJSOM = (inputData, opts = {}) => {
  if (!inputData) return inputData;
  const { expanded, groupBy } = opts;
  const getValues = spObject => {
    if (!spObject) return;
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

export const prepareResponseREST = data => {
  let result;
  if (typeOf(data) === 'array') {
    let outputData = [];
    for (let elementData of data) {
      try {
        result = JSON.parse(elementData.body).d;
      } catch (err) {
        outputData.push(elementData)
      }
      if (result) {
        if (result.results) {
          outputData = [...outputData, ...result.results]
        } else {
          outputData.push(result)
        }
      }
    }
    return outputData;
  } else {
    try {
      result = JSON.parse(data.body).d;
    } catch (err) {
      return data
    }
    if (result) {
      return result.results ? result.results : result;
    }
  }
}

export const getChunkedArray = (elements, size) => {
  let chunks = [];
  let chunkSize = size || this.REQUEST_BUNDLE_MAX_SIZE;
  if (elements === void 0) return [chunks];
  if (typeOf(elements) === 'array') {
    elements = [...elements];
    while (elements.length) chunks.push(elements.splice(0, chunkSize))
  } else {
    chunks.push([elements])
  }
  return chunks;
}

export const getRequestDigestREST = async relPath => {
  return (await axios.post(`${(!relPath || relPath === '/' ? '' : relPath)}/_api/contextinfo`, null, {
    headers: { 'Accept': 'application/json; odata=verbose' }
  })).data.d.GetContextWebInformation.FormDigestValue;
}

export const getClientContext = contextUrl => {
  const clientContext = new SP.ClientContext(contextUrl);
  clientContext.set_requestTimeout(REQUEST_TIMEOUT);
  return clientContext;
}

export const load = (clientContext, data, opts = {}) => {
  let { groupBy, view } = opts;
  if (view) {
    view = typeOf(view) === 'array' ? view : [view]
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
      !opts.silent && console.error(args.get_message());
      reject(args)
    })
  })

export const execute = async (elementName, element, handler, opts = {}) =>
  this[`_execute${opts.rest ? 'REST' : 'JSOM'}`](elementName, element, handler, opts)

export const _executeREST = async (elementName, element, handler, opts = {}) => {
  let {
    silent,
    actionType,
    logTextToAdd = '',
    cached,
    getter
  } = opts;
  let restObjectsChunks = [];
  let elementNameREST = `${elementName}REST`;
  let pathBuilder = new PathBuilder;
  pathBuilder.set('clientContext', this.contextUrl);
  pathBuilder.set('parentElement', this.parent.element);
  await this.iterateChunks(element, async (elementsChunk) => {
    let elementUrl;
    let elementUrls = [];
    let restObjects = [];
    for (let elementChunk of elementsChunk) {
      let {
        request,
        params = {}
      } = await handler(elementChunk);
      if (!request) continue;
      elementUrl = typeOf(elementChunk) === 'object' ? (elementChunk.Url || elementChunk.Title) : elementChunk;
      elementUrls.push(elementUrl);
      pathBuilder.set('element', elementUrl);
      if (cached && actionType !== 'delete' && actionType !== 'recycle' && (!actionType || actionType === 'get' || actionType === 'create')) {
        try {
          let restObjectCached = cache(pathBuilder.build(), elementNameREST).get();
          if (getter) restObjectCached = getter(restObjectCached);
          restObjects.push(restObjectCached);
          continue;
        } catch (err) { }
      }
      let httpProvider = params.httpProvider || this.requestExecutor.bind(this);
      let restObject = await httpProvider(request);
      if (params.errorHandler) params.errorHandler(restObject);
      if (getter) restObject = getter(restObject);
      cache(pathBuilder.build(), elementNameREST, restObject)[actionType === 'delete' || actionType === 'recycle' ? 'delete' : 'set']();
      restObjects.push(restObject);
    }
    restObjectsChunks = [...restObjectsChunks, ...restObjects];
    (this.debug || !silent) && (actionType && actionType !== 'get') &&
      console.log(`${this.ACTION_TYPES[actionType]} ${elementName} at ${this.contextUrl}: ${elementUrls.join(', ')}${logTextToAdd}`)
  }, actionType)
  if (actionType !== 'delete' && actionType !== 'recycle') return this.prepareResponseREST(typeOf(element) === 'array' ? restObjectsChunks : restObjectsChunks[0], opts);
}

export const _executeJSOM = async (elementName, element, handler, opts = {}) => {
  let {
    view,
    cached,
    actionType,
    silent,
    getter,
    logTextToAdd = ''
  } = opts;
  let pathBuilder = new PathBuilder;
  pathBuilder.set('clientContext', this.contextUrl);
  pathBuilder.set('parentElement', this.parent.element);
  let spObjectChunks = [];
  let isNotDeleteAction = actionType !== 'delete' && actionType !== 'recycle';
  await this.iterateChunks(element, async elementsChunk => {
    let elementUrl;
    let elementUrls = [];
    let spObjects = [];
    let clientContextsMap = new Map;
    this._initClientContext();
    for (let elementChunk of elementsChunk) {
      let spObject = await handler(elementChunk);
      if (!spObject) continue;
      if (typeOf(spObject) === 'array') {
        spObjects = [...spObjects, ...spObject];
        continue;
      }
      if (spObject.cacheUrl) {
        elementUrl = spObject.cacheUrl;
      } else if (typeOf(elementChunk) === 'object') {
        elementUrl = elementChunk.Url || elementChunk.Title;
      } else {
        elementUrl = elementChunk;
      }
      pathBuilder.set('element', elementUrl);
      let clientContext = spObject.get_context();
      let currentContextUrl = clientContext.get_url();
      if (isNotDeleteAction) this.load(clientContext, spObject, view);
      if (cached && isNotDeleteAction && (!actionType || actionType === 'get' || actionType === 'create')) {
        try {
          let spObjectCached = cache(pathBuilder.build(), elementName, spObject).get();
          if (getter) spObjectCached = getter(spObjectCached);
          spObjects.push(spObjectCached);
          continue;
        } catch (err) { }
      };
      !clientContextsMap.has(currentContextUrl) && clientContextsMap.set(currentContextUrl, {
        clientContext,
        spObjects: []
      })
      clientContextsMap.get(currentContextUrl).spObjects.push({
        spObject,
        elementUrl
      })
    }
    for (let [currentContextUrl, elementSP] of clientContextsMap) {
      await this.executeQueryAsync(elementSP.clientContext, opts);
      for (let {
        spObject,
        elementUrl
      } of elementSP.spObjects) {
        if (spObject.get_id && isNotDeleteAction) elementUrl = spObject.get_id();
        pathBuilder.set('element', elementUrl);
        elementUrls.push(elementUrl);
        cache(pathBuilder.build(), elementName, spObject)[isNotDeleteAction ? 'set' : 'delete']();
        if (getter) spObject = getter(spObject);
        spObjects.push(spObject);
      }
    }
    spObjectChunks = [...spObjectChunks, ...spObjects];
    let elementsToLog = elementUrls.join(', ');
    elementsToLog && (this.debug || !silent) && (actionType && actionType !== 'get') &&
      console.log(`${this.ACTION_TYPES[actionType]} ${elementName} at ${this.contextUrl}: ${elementsToLog}${logTextToAdd}`);
  }, actionType)
  if (isNotDeleteAction) {
    let isElementArray = typeOf(element) === 'array';
    return this.prepareResponseJSOM(isElementArray || (!isElementArray && spObjectChunks.length > 1) ? spObjectChunks : spObjectChunks[0], opts);
  }
}

export const iterateChunks = async (elements, handlerAsync, actionType) => {
  let elementsType = typeOf(elements);
  if (!actionType ||
    actionType === 'get' ||
    elementsType !== 'array') {
    await handlerAsync(elementsType !== 'array' ? [elements] : elements);
  } else {
    let chunkedElements = this.getChunkedArray(elements);
    for (let chunk of chunkedElements) await handlerAsync(chunk);
  }
}

export const get_ = async (methodName, opts = {}) => {
  let {
    args = []
  } = opts;
  return this.executeBinded(element => {
    if (element) {
      return this._getSPObject(element);
    } else {
      console.log(this._getSPObject(element))
    }
  }, {
      ...opts,
      raw: true,
      getter: spObject => spObject[this.getValidMethodName(spObject, methodName)](...args)
    })
}

export const requestExecutor = async (params = {}) =>
  new Promise((resolve, reject) => {
    new SP.RequestExecutor(this.contextUrl).executeAsync({
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
    let bytes = new Uint8Array(data);
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
      typeOf(valueInt) === 'number' && valueInt > 0 && lookupValue.set_lookupId(valueInt);
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
          view: ['TypeAsString', 'InternalName', 'Title'],
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
