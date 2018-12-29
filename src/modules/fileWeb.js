import * as cache from './../cache';
import * as utility from './../utility';
import spx from './../modules/site';

export default class FileWeb {
  constructor(parent, elementUrl) {
    this._elementUrl = elementUrl;
    this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
    this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl];
    this._parent = parent;
    this._contextUrlIsArray = this._parent._contextUrlIsArray;
    this._contextUrls = this._parent._contextUrls;
  }

  // Inteface

  async get(opts = {}) {
    if (opts.blob) {
      return this._executeREST(null, spObjectUrl => ({
        request: {
          url: `${spObjectUrl}/$value`,
          binaryStringResponseBody: true
        }
      }), opts)
    } else {
      return this._execute(null, spObject => (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
    }
  }

  async create(opts = {}) {
    return this._execute('create', async (spContextObject, elementUrl) => {
      const {
        Url,
        Content = '',
        Overwrite = true
      } = elementUrl;

      const { folder, filename } = utility.getFolderAndFilenameFromUrl(Url);
      if (folder) await spx(spContextObject.get_context().get_url()).folder(folder).create({ silent: true, expanded: true, view: ['Name'] }).catch(() => { });
      const fileCreationInfo = new SP.FileCreationInformation;
      utility.setFields(fileCreationInfo, {
        set_url: filename,
        set_content: utility.convertFileContent(Content),
        set_overwrite: Overwrite
      })
      const spObject = spContextObject.add(fileCreationInfo);
      spObject.cachePath = 'property';
      return spObject;
    }, opts)
  }

  async update(opts = {}) {
    return this._execute('update', async (spObject, elementUrl) => {
      const { Content } = elementUrl;
      const binaryInfo = new SP.FileSaveBinaryInformation;
      if (Content !== void 0) binaryInfo.set_content(utility.convertFileContent(Content));
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
    return this._execute('copy', async (spObject, elementUrl) => {
      spObject.copyTo(elementUrl.To.replace(/^\//, ''));
      spObject.cachePath = 'property';
      return spObject;
    }, opts);
  }

  async move(opts = {}) {
    return this._execute('move', async (spObject, elementUrl) => {
      spObject.moveTo(elementUrl.To.replace(/^\//, ''));
      spObject.cachePath = 'property';
      return spObject;
    }, opts);
  }

  // Internal

  get _name() { return 'file' }

  async _execute(actionType, spObjectGetter, opts = {}) {
    const { cached } = opts;
    let isArrayCounter = 0;
    const clientContexts = {};
    if (opts.asItem) opts.view = ['ListItemAllFields'];
    const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
      let needToQuery;
      let totalElements = 0;
      let clientContext = utility.getClientContext(contextUrl);
      const spObjectsToCache = new Map;
      const contextUrls = contextUrl.split('/');
      clientContexts[contextUrl] = [clientContext];
      const spObjects = await Promise.all(this._elementUrls.map(async elementUrl => {
        const file = this._liftElementUrlType(elementUrl);
        let fileUrl = file.Url;
        if (actionType === 'create') fileUrl = fileUrl.split('/').slice(0, -1).join('/') + '/';
        if (actionType && ++totalElements >= utility.REQUEST_BUNDLE_MAX_SIZE) {
          clientContext = utility.getClientContext(contextUrl);
          clientContexts[contextUrl].push(clientContext);
          totalElements = 0;
        }

        const spObject = await spObjectGetter(this._getSPObject(clientContext, fileUrl), file);
        !!spObject.getEnumerator && isArrayCounter++;
        const fileUrls = fileUrl.split('/');
        const cachePaths = [...contextUrls, ...fileUrls, this._name, spObject.cachePath];
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
            spObjectsToCache.set(cachePaths, currentSPObjects)
            return currentSPObjects;
          }
        }
      }))

      if (needToQuery) {
        await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
          contextAcc.concat(clientContexts[contextUrl].map(clientContext => utility.executeQueryAsync(clientContext, opts))), []));
        spObjectsToCache.forEach((value, key) => cache.set(key, value))
      };
      return spObjects;
    }))

    this._log(actionType, opts);
    opts.isArray = isArrayCounter || this._contextUrlIsArray || this._elementUrlIsArray;
    return utility.prepareResponseJSOM(elements, opts);
  }

  async _executeREST(actionType, restObjectGetter, opts = {}) {
    const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
      contextAcc.concat(this._elementUrls.map(async elementUrl => {
        const restObject = await restObjectGetter(this._getRESTObject(contextUrl, elementUrl), elementUrl);
        const { request, params = {} } = restObject;
        const httpProvider = params.httpProvider || utility.requestExecutor.bind(null, contextUrl);
        return httpProvider(request);
      })), []));
    this._log(actionType, opts);
    opts.isArray = this._contextUrlIsArray || this._elementUrlIsArray;
    return utility.prepareResponseREST(elements, opts);
  }

  _getRESTObject(contextUrl, elementUrl) {
    const { folder, filename } = utility.getFolderAndFilenameFromUrl(elementUrl);
    const url = `${contextUrl}/_api/web`;
    if (elementUrl) {
      const fullUrl = folder ? `${contextUrl}/${folder}` : contextUrl;
      return filename ?
        `${url}/getfilebyserverrelativeurl(${fullUrl}/${filename})`.replace(/\/\/+/g, '/') :
        `${url}/getfolderbyserverrelativeurl('${fullUrl}')/files`
    } else {
      return `${url}/getrootfolder/files`
    }
  }

  _getSPObject(clientContext, elementUrl) {
    const spContextObject = this._parent._getSPObject(clientContext);
    const { folder, filename } = utility.getFolderAndFilenameFromUrl(elementUrl);
    const contextUrl = clientContext.get_url();
    if (elementUrl) {
      const fullUrl = folder ? `${contextUrl}/${folder}` : contextUrl;
      return filename ?
        spContextObject.getFileByServerRelativeUrl(`${fullUrl}/${filename}`.replace(/\/\/+/g, '/')) :
        spContextObject.getFolderByServerRelativeUrl(`${fullUrl}`).get_files()
    } else {
      return spContextObject.get_rootFolder().get_files();
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
        }).join(', ')}`);
  }
}