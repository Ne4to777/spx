import {
  ContextUrlBox,
  getClientContext,
  prepareResponseJSOM,
  substitutionI,
  executeJSOM,
  pipe,
  identity,
  methodEmpty,
  methodI,
  getInstance,
  getContext,
  getWeb
} from './../utility'
import * as cache from './../cache'
import site from './../modules/web'
import recycleBin from './../modules/recycleBin'
import user from './../modules/user'
// import tag from './../modules/tag'

// Internal

const NAME = 'site';

const getSPObject = methodEmpty('get_site');
const getListTemplates = methodI('getCustomListTemplates');

const box = getInstance(ContextUrlBox)('/');

const execute = cacheLeaf => spObjectGetter => async (opts = {}) => {
  const { cached } = opts;
  const clientContext = getClientContext('/');
  const spObject = spObjectGetter(getSPObject(clientContext));
  const cachePath = [NAME, cacheLeaf + 'Collection'];
  const spObjectCached = cached ? cache.get(cachePath) : null;
  if (cached && spObjectCached) return prepareResponseJSOM(opts)(currentSPObjects);
  const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts);
  cache.set(currentSPObjects)(cachePath);
  return prepareResponseJSOM(opts)(currentSPObjects);
}

// Interface

// site.tag = tag;
site.user = user;
site.recycleBin = recycleBin({ box, getSPObject });

site.get = execute('properties')(identity);
site.getCustomListTemplates = execute('customListTemplates')(substitutionI(pipe([getContext, getWeb]))(getListTemplates));
site.getWebTemplates = execute('webTemplates')(spParentObject => spParentObject.getWebTemplates(1033, 0));

export default site;