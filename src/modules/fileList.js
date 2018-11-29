import axios from 'axios';
import * as cache from './../cache';
import * as utility from './../utility';

export default class FileList {
  constructor(parent, elementUrl) {
    this._elementUrl = elementUrl;
    this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
    this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl];
    this._parent = parent;
    this._listUrlIsArray = this._parent._elementUrlIsArray;
    this._listUrls = this._parent._elementUrls;
    this._contextUrlIsArray = this._parent._contextUrlIsArray;
    this._contextUrls = this._parent._contextUrls;
  }

  // Inteface

  async get(opts = {}) {
    if (opts.blob) {
      return this._executeREST(null, (contextUrl, listUrl, elementUrl) => ({
        request: {
          url: `${this._getRESTObject(contextUrl, listUrl, elementUrl)}/$value`,
          binaryStringResponseBody: true
        }
      }), opts)
    } else {
      return this._execute(null, spObject => (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
    }
  }

  async create(opts) {
    let contentIsString;
    for (let contextUrl of this._contextUrls) {
      if (contentIsString) break;
      for (let listUrl of this._listUrls) {
        if (contentIsString) break;
        for (let elementUrl of this._elementUrls) {
          if (elementUrl && typeOf(elementUrl.Content) === 'string') {
            contentIsString = true;
            break
          }
        }
      }
    }
    return contentIsString ?
      this._execute('create', (spContextObject, listUrl, elementUrl) => this._createWithJSOM(spContextObject, listUrl, elementUrl), opts) :
      this._executeREST('create', (contextUrl, listUrl, elementUrl) => this._createWithRESTFromFile(contextUrl, listUrl, elementUrl), opts)
  }

  async update(opts = {}) {
    return this._execute('update', async (spObject, listUrl, elementUrl) => {
      const { Content } = elementUrl;
      const fieldsToCreate = {};
      for (let fieldName in elementUrl) {
        if (fieldName === 'Content' || fieldName === 'Url') continue;
        const field = elementUrl[fieldName];
        fieldsToCreate[fieldName] = typeOf(field) === 'array' ? field.join(';#;#') : field;
      }
      const binaryInfo = new SP.FileSaveBinaryInformation;
      if (Content !== void 0) binaryInfo.set_content(utility.convertFileContent(Content));
      binaryInfo.set_fieldValues(fieldsToCreate);
      spObject.saveBinary(binaryInfo);
      spObject.cachePath = 'property';
      return spObject
    }, opts);
  }

  async delete(opts = {}) {
    const { noRecycle } = opts;
    return this._execute(noRecycle ? 'delete' : 'recycle', spObject => {
      spObject[noRecycle ? 'deleteObject' : 'recycle']();
      spObject.cachePath = 'property';
      return spObject;
    }, opts)
  }

  async copy(opts = {}) {
    return this._execute('copy', (spObject, listUrl, elementUrl) => {
      spObject.copyTo(`${listUrl}/${elementUrl.To}`.replace(/\/\/+/g, '/'));
      spObject.cachePath = 'property';
      return spObject;
    }, opts);
  }

  async move(opts = {}) {
    return this._execute('move', (spObject, listUrl, elementUrl) => {
      spObject.moveTo(`${listUrl}/${elementUrl.To}`.replace(/\/\/+/g, '/'));
      spObject.cachePath = 'property';
      return spObject;
    }, opts);
  }
  // Internal

  get _name() { return 'file' }

  async _createWithRESTFromFile(contextUrl, listUrl, elementUrl) {
    let founds;
    const inputs = [];
    const { Url, Content = '', Overwrite = true, OnProgress = () => { } } = elementUrl;
    const { folder = '', filename } = utility.getFolderAndFilenameFromUrl(Url);
    if (folder) await spx(contextUrl).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(() => { });
    const requiredInputs = {
      __REQUESTDIGEST: true,
      __VIEWSTATE: true,
      __EVENTTARGET: true,
      __EVENTVALIDATION: true,
      ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation: true,
      ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle: true,
    }
    const listGUID = (await spx(contextUrl).list(listUrl).get({ cached: true, view: 'Id' })).Id.toString();
    const res = await axios.get(`${contextUrl}/_layouts/15/Upload.aspx?List={${listGUID}}`);
    const formMatches = res.data.match(/<form(\w|\W)*<\/form>/);
    const inputRE = /<input[^<]*\/>/g;
    while (founds = inputRE.exec(formMatches)) {
      let item = founds[0];
      const id = item.match(/id=\"([^\"]+)\"/)[1];
      if (requiredInputs[id]) {
        switch (id) {
          case '__EVENTTARGET':
            item = item.replace(/value="[^\"]*"/, 'value="ctl00$PlaceHolderMain$ctl03$RptControls$btnOK"')
            break;
          case 'ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation':
            item = item.replace(/value="[^\"]*"/, `value="/${folder.replace(/^\//, '')}"`);
            break;
          case 'ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle':
            if (!Overwrite) item = item.replace(/checked="[^\"]*"/, '');
            break;
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

  async _createWithJSOM(spContextObject, listUrl, elementUrl) {
    const { Url, Content = '', Overwrite = true } = elementUrl;
    const { folder, filename } = utility.getFolderAndFilenameFromUrl(Url);
    if (folder) await spx(spContextObject.get_context().get_url()).list(listUrl).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(() => { });
    const fileCreationInfo = new SP.FileCreationInformation;
    utility.setFields(fileCreationInfo, {
      set_url: filename,
      set_content: '',
      set_overwrite: Overwrite
    })
    const fieldsToCreate = {};
    for (let fieldName in elementUrl) {
      if (fieldName === 'Content' || fieldName === 'Url' || fieldName === 'Overwrite') continue;
      const field = elementUrl[fieldName];
      fieldsToCreate[fieldName] = typeOf(field) === 'array' ? field.join(';#;#') : field;
    }
    const binaryInfo = new SP.FileSaveBinaryInformation;
    utility.setFields(binaryInfo, {
      set_content: utility.convertFileContent(Content),
      set_fieldValues: fieldsToCreate
    })
    const spObject = spContextObject.add(fileCreationInfo);
    spObject.saveBinary(binaryInfo);
    spObject.cachePath = 'property';
    return spObject;
  }

  async _execute(actionType, spObjectGetter, opts = {}) {
    let needToQuery;
    let isArrayCounter = 0;
    const clientContexts = {};
    const spObjectsToCache = new Map;
    const { cached, parallelized = actionType !== 'create' } = opts;
    opts.view = opts.view || (actionType ? ['ServerRelativeUrl'] : void 0);
    const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) => {
      let totalElements = 0;
      const contextUrls = contextUrl.split('/');
      let clientContext = utility.getClientContext(contextUrl);
      clientContexts[contextUrl] = [clientContext];
      return contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
        listAcc.concat(this._elementUrls.map(async elementUrl => {
          const file = this._liftElementUrlType(elementUrl);
          let fileUrl = file.Url;
          if (actionType === 'create') fileUrl = fileUrl.split('/').slice(0, -1).join('/') + '/';
          if (actionType && ++totalElements >= utility.REQUEST_BUNDLE_MAX_SIZE) {
            clientContext = utility.getClientContext(contextUrl);
            clientContexts[contextUrl].push(clientContext);
            totalElements = 0;
          }
          const spObject = await spObjectGetter(this._getSPObject(clientContext, listUrl, fileUrl), listUrl, elementUrl);
          !!spObject.getEnumerator && isArrayCounter++;
          const fileUrls = fileUrl.split('/');
          const cachePaths = [...contextUrls, listUrl, ...fileUrls, this._name, spObject.cachePath];
          utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(cachePaths.slice(0, -1));
          if (actionType === 'delete' || actionType === 'recycle') {
            needToQuery = true;
          } else {
            const spObjectCached = cached ? cache.get(cachePaths) : null;
            if (cached && spObjectCached) {
              return spObjectCached;
            } else {
              needToQuery = true;
              const currentSPObjects = utility.load(clientContext, spObject, opts);
              !actionType && spObjectsToCache.set(cachePaths, currentSPObjects)
              return currentSPObjects;
            }
          }
        })), []))
    }, []))

    if (needToQuery) {
      await Promise.all(parallelized ?
        this._contextUrls.reduce((contextAcc, contextUrl) =>
          contextAcc.concat(clientContexts[contextUrl].map(clientContext => utility.executeQueryAsync(clientContext, opts))), []) :
        this._contextUrls.map(async (contextUrl) => {
          for (let clientContext of clientContexts[contextUrl]) await utility.executeQueryAsync(clientContext, opts)
        }));
      !actionType && spObjectsToCache.forEach((value, key) => cache.set(key, value))
    };
    this._log(actionType, opts);
    opts.isArray = isArrayCounter || this._contextUrlIsArray || this._listUrlIsArray || this._elementUrlIsArray;
    return utility.prepareResponseJSOM(elements, opts);
  }

  async _executeREST(actionType, restObjectGetter, opts = {}) {
    const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
      contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
        listAcc.concat(this._elementUrls.map(async elementUrl => {
          const restObject = await restObjectGetter(contextUrl, listUrl, elementUrl);
          const { request, params = {} } = restObject;
          const httpProvider = params.httpProvider || utility.requestExecutor.bind(null, contextUrl);
          return httpProvider(request);
        })), [])), []));
    this._log(actionType, opts);
    opts.isArray = this._contextUrlIsArray || this._listUrlIsArray || this._elementUrlIsArray;
    return utility.prepareResponseREST(elements, opts);
  }

  _getRESTObject(contextUrl, listUrl, elementUrl = '') {
    const { folder = '', filename } = utility.getFolderAndFilenameFromUrl(elementUrl);
    const folderTrimmed = folder.replace(/^\//, '');
    const rootFolder = `${contextUrl}/_api/web/lists/${utility.isGUID(listUrl) ? `getbyid(\'${listUrl}\')` : `getbytitle(\'${listUrl}\')`}/rootfolder`;
    if (filename) {
      return folder ?
        `${rootFolder}/folders/getbyurl('${folderTrimmed}')/files('${filename}')` :
        `${rootFolder}/files('${filename}')`
    } else {
      return folder ?
        `${rootFolder}/folders/getbyurl('${folderTrimmed}')/files` :
        `${rootFolder}/files`
    }
  }

  _getSPObject(clientContext, listUrl, elementUrl) {
    const spContextObject = this._parent._getSPObject(clientContext, listUrl);
    const spFolderObject = spContextObject.get_rootFolder();
    if (elementUrl) {
      const { folderUrl, filename } = utility.getFolderAndFilenameFromUrl(elementUrl);
      if (folderUrl) {
        const files = utility.getSPFolderByUrl(spFolderObject, folderUrl).get_files();
        return filename ? files.getByUrl(filename) : files;
      } else {
        if (/\/$/.test(folderUrl)) {
          return utility.getSPFolderByUrl(spFolderObject, folderUrl).get_files();
        } else {
          const files = spFolderObject.get_files();
          return filename ? files.getByUrl(filename) : files;
        }
      }
    } else {
      return spContextObject.get_files();
    }
  }

  _liftElementUrlType(elementUrl) {
    switch (typeOf(elementUrl)) {
      case 'object':
        return elementUrl;
      case 'string':
      default:
        return { Url: elementUrl || '/' }
    }
  }

  _log(actionType, opts = {}) {
    !opts.silent && actionType &&
      console.log(`${
        utility.ACTION_TYPES[actionType]} ${
        this._contextUrls.length * this._elementUrls.length} ${
        this._name}(s) at ${
        this._contextUrls.join(', ')}: ${
        this._elementUrls.map(el => {
          const element = this._liftElementUrlType(el);
          return element.To ? `${element.Url} -> ${element.To}` : element.Url
        }).join(', ')} in ${
        this._listUrls.join(', ')}`);
  }
}