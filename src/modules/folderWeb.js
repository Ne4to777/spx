import * as utility from './../utility';
import * as cache from './../cache';
import spx from './../modules/site';

export default class FolderWeb {
  constructor(parent, elementUrl) {
    this._elementUrl = elementUrl;
    this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
    this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl];
    this._parent = parent;
    this._contextUrlIsArray = this._parent._contextUrlIsArray;
    this._contextUrls = this._parent._contextUrls;
  }

  // Inteface

  async get(opts) {
    return this._execute(null, spObject => (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
  }

  async create(opts) {
    return this._execute('create', async (spContextObject, elementUrl) => {
      if (!elementUrl) throw new Error('no folders to create');
      const clientContext = spContextObject.get_context();
      const elementSplits = elementUrl.split('/');
      const lastFolderName = elementSplits.pop();
      const parentFolderUrl = elementSplits.join('/');
      if (parentFolderUrl) await spx(clientContext.get_url()).folder(parentFolderUrl).create({ silent: true, expanded: true, view: ['Name'] }).catch(() => { });
      const folders = this._getSPObject(clientContext, `${parentFolderUrl}/`);
      const spObject = folders.add(lastFolderName);
      spObject.cachePath = 'property';
      return spObject
    }, opts)
  }

  async delete(opts = {}) {
    return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject =>
      (spObject.recycle && spObject[opts.noRecycle ? 'deleteObject' : 'recycle'](), spObject.cachePath = 'property', spObject), opts)
  }

  // Internal

  get _name() { return 'folder' }

  async _execute(actionType, spObjectGetter, opts = {}) {
    const { cached } = opts;
    let isArrayCounter = 0;
    const clientContexts = {};
    const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
      let needToQuery;
      let totalElements = 0;
      let clientContext = utility.getClientContext(contextUrl);
      const spObjectsToCache = new Map;
      const contextUrls = contextUrl.split('/');
      clientContexts[contextUrl] = [clientContext];
      const spObjects = await Promise.all(this._elementUrls.map(async elementUrl => {
        let folderUrl = elementUrl || '/';
        if (actionType === 'create') folderUrl = '';
        if (actionType && ++totalElements >= utility.REQUEST_BUNDLE_MAX_SIZE) {
          clientContext = utility.getClientContext(contextUrl);
          clientContexts[contextUrl].push(clientContext);
          totalElements = 0;
        }
        const spObject = await spObjectGetter(this._getSPObject(clientContext, folderUrl), elementUrl);
        !!spObject.getEnumerator && isArrayCounter++;
        const folderUrls = folderUrl.split('/');
        const cachePaths = [...contextUrls, ...folderUrls, this._name, spObject.cachePath];
        utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(cachePaths.slice(0, -folderUrls.length));
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

  _getSPObject(clientContext, elementUrl) {
    if (clientContext) {
      const folder = this._parent._getSPObject(clientContext).getFolderByServerRelativeUrl(elementUrl.replace(/\/$/, ''));
      return /\/$/.test(elementUrl) ? folder.get_folders() : folder
    } else {
      return this._parent._getSPObject().get_rootFolder().get_folders();
    }
  }

  _log(actionType, opts = {}) {
    !opts.silent && actionType &&
      console.log(`${
        utility.ACTION_TYPES[actionType]} ${
        this._contextUrls.length * this._elementUrls.length} ${
        this._name}(s) at ${
        this._contextUrls.join(', ')}: ${
        this._elementUrls.join(', ')}`);
  }
}