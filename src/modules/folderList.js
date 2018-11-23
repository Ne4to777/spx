import * as utility from './../utility';
import * as cache from './../cache';
import spx from './../modules/site';

const _fieldsInfo = {};

export default class FolderList {
  constructor(parent, elementUrl) {
    this._elementUrl = elementUrl;
    this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
    this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl];
    this._parent = parent;
    this._contextUrlIsArray = this._parent._contextUrlIsArray;
    this._contextUrls = this._parent._contextUrls;
    this._listUrlIsArray = this._parent._elementUrlIsArray;
    this._listUrls = this._parent._elementUrls;
  }

  // Inteface

  async get(opts) {
    return this._execute(null, spObject =>
      (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
  }

  async create(opts) {
    await utility.initFieldsInfo(this._contextUrls, this._listUrls, this._fieldsInfo);
    return this._execute('create', async (spContextObject, listUrl, element) => {
      if (!element) throw new Error('no folders to create');
      const elementNew = Object.assign({}, element);
      const url = element.Url;
      delete elementNew.Url;
      const clientContext = spContextObject.get_context();
      const itemCreationInfo = new SP.ListItemCreationInformation;
      const parentFolderUrl = url.split('/').slice(0, -1);
      const contextUrl = clientContext.get_url();
      parentFolderUrl && await spx(contextUrl).list(listUrl).folder(parentFolderUrl).create({
        view: ['FileLeafRef'],
        silent: true,
        expanded: true,
      }).catch(() => { });
      itemCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
      itemCreationInfo.set_leafName(url);
      const spObject = spContextObject.addItem(itemCreationInfo);
      await utility.setItem(_fieldsInfo[contextUrl][listUrl], spObject, elementNew);
      spObject.update()
      spObject.cachePath = 'property';
      return spObject
    }, opts);
  }

  async delete(opts = {}) {
    return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject => {
      spObject.recycle && spObject[opts.noRecycle ? 'deleteObject' : 'recycle']();
      spObject.cachePath = 'property';
      return spObject;
    }, opts)
  }

  // Internal

  get _name() { return 'folder' }

  async _execute(actionType, spObjectGetter, opts = {}) {
    const { cached, parallelized = actionType !== 'create' } = opts;
    opts.view = opts.view || (actionType ? ['ID'] : void 0);
    let isArrayCounter = 0;
    const clientContexts = {};
    const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
      let needToQuery;
      let totalElements = 0;
      const spObjects = [];
      const spObjectsToCache = new Map;
      const contextUrls = contextUrl.split('/');
      clientContexts[contextUrl] = {};
      for (let listUrl of this._listUrls) {
        let clientContext = utility.getClientContext(contextUrl);
        clientContexts[contextUrl][listUrl] = [clientContext]
        for (let elementUrl of this._elementUrls) {
          const element = this._liftElementUrlType(elementUrl);
          let folderUrl = element.Url;
          if (actionType === 'create') folderUrl = '';
          if (actionType && ++totalElements >= utility.REQUEST_BUNDLE_MAX_SIZE) {
            clientContext = utility.getClientContext(contextUrl);
            clientContexts[contextUrl][listUrl].push(clientContext);
            totalElements = 0;
          }
          const spObject = await spObjectGetter(this._getSPObject(clientContext, listUrl, folderUrl), listUrl, element);
          !!spObject.getEnumerator && isArrayCounter++;
          const folderUrls = folderUrl.split('/');
          const cachePaths = [...contextUrls, listUrl, ...folderUrls, this._name, spObject.cachePath];
          utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(cachePaths.slice(0, -folderUrls.length));
          if (actionType === 'delete' || actionType === 'recycle') {
            needToQuery = true;
          } else {
            const spObjectCached = cached ? cache.get(cachePaths) : null;
            if (cached && spObjectCached) {
              spObjects.push(spObjectCached);
            } else {
              needToQuery = true;
              const currentSPObjects = utility.load(clientContext, spObject, opts);
              spObjectsToCache.set(cachePaths, currentSPObjects)
              spObjects.push(currentSPObjects);
            }
          }
        }
      }
      if (needToQuery) {
        if (parallelized) {
          await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
            contextAcc.concat(this._listUrls.reduce((listAcc, listUrl) =>
              listAcc.concat(clientContexts[contextUrl][listUrl].map(clientContext => utility.executeQueryAsync(clientContext, opts))), [])), []));
        } else {
          await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
            contextAcc.concat(this._listUrls.map(async listUrl => {
              for (let clientContext of clientContexts[contextUrl][listUrl]) await utility.executeQueryAsync(clientContext, opts)
            })), []));
        }
        spObjectsToCache.forEach((value, key) => cache.set(key, value))
      };
      return spObjects;
    }))

    this._log(actionType, opts);
    opts.isArray = isArrayCounter || this._contextUrlIsArray || this._listUrlIsArray || this._elementUrlIsArray;
    return utility.prepareResponseJSOM(elements, opts);
  }

  _getSPObject(clientContext, listUrl, elementUrl) {
    if (clientContext) {
      if (elementUrl) {
        const folderUrls = elementUrl.split('/').filter(el => !!el.length);
        const getSPFolderR = (base, i) => {
          if (i < folderUrls.length) {
            return getSPFolderR(base.get_folders().getByUrl(folderUrls[i]), ++i)
          } else {
            return /\/$/.test(elementUrl) ? base.get_folders() : base;
          }
        }
        return getSPFolderR(this._parent._getSPObject(clientContext, listUrl).get_rootFolder(), 0);
      } else {
        return this._parent._getSPObject(clientContext, listUrl)
      }
    } else {
      throw new Error('ClientContext is missed')
    }
  }

  get _fieldsInfo() { return _fieldsInfo }

  _liftElementUrlType(elementUrl) {
    switch (typeOf(elementUrl)) {
      case 'object':
        if (!elementUrl.Url) elementUrl.Url = elementUrl.Title;
        if (!elementUrl.Title) elementUrl.Title = utility.getLastPath(elementUrl.Url);
        return elementUrl;
      case 'string':
        elementUrl = elementUrl.replace(/\/$/, '');
        return {
          Url: elementUrl,
          Title: utility.getLastPath(elementUrl)
        }
    }
  }

  _log(actionType, opts = {}) {
    !opts.silent && actionType &&
      console.log(`${
        utility.ACTION_TYPES[actionType]} ${
        this._contextUrls.length * this._listUrls.length * this._elementUrls.length} ${
        this._name}(s) at ${
        this._contextUrls.join(', ')} in ${
        this._listUrls.join(', ')}`);
  }
}