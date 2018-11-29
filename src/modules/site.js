import Web from './../modules/web';
import RecycleBin from './../modules/recycleBin';
import * as cache from './../cache';
import * as utility from './../utility';

const site = contextUrl => new Web(contextUrl);

// Interface

site.get = async opts => execute(spObject => (spObject.cachePath = 'properties', spObject), opts);

site.getCustomListTemplates = async opts =>
  execute(spContextObject => {
    const spObject = spContextObject.getCustomListTemplates(spContextObject.get_context().get_web());
    spObject.cachePath = 'customListTemplates';
    return spObject
  }, opts);

// Internal

site._name = 'site';

site._contextUrls = ['/'];

const execute = async (spObjectGetter, opts = {}) => {
  const { cached } = opts;
  const clientContext = utility.getClientContext('/');
  const spObject = spObjectGetter(clientContext.get_site());
  const cachePaths = ['site', spObject.cachePath];
  const spObjectCached = cached ? cache.get(cachePaths) : null;
  if (cached && spObjectCached) {
    return utility.prepareResponseJSOM(spObjectCached, opts)
  } else {
    const spObjects = utility.load(clientContext, spObject, opts);
    await utility.executeQueryAsync(clientContext, opts);
    cache.set(cachePaths, spObjects);
    return utility.prepareResponseJSOM(spObjects, opts)
  }
}

site._getSPObject = clientContext => clientContext.get_site();

site.recycleBin = new RecycleBin(site);
window.spx = site;
export default site