import {
	ACTION_TYPES,
	ACTION_TYPES_TO_UNSET,
	ContextUrlBox,
	getClientContext,
	prepareResponseJSOM,
	setFields,
	popSlash,
	pipe,
	slice,
	methodEmpty,
	getInstance,
	overstep,
	executeJSOM,
	urlSplit,
	executorJSOM,
	getInstanceEmpty,
	methodI,
	getParentUrl,
	prop,
	isStringEmpty,
	hasUrlTailSlash
} from './../utility'
import * as cache from './../cache'

// import search from './../modules/search'
import list from './../modules/list'
import folder from './../modules/folderWeb'
import file from './../modules/fileWeb'
import recycleBin from './../modules/recycleBin'

// Internal

const APP_WEB_TEMPLATE = '{E2A30D74-39CB-429E-A5E0-4C775BE848CE}#Default'
const NAME = 'web';

const getSPObject = methodEmpty('get_web');
const getSPObjectCollection = pipe([getSPObject, methodEmpty('get_webs')]);

const report = ({ silent, actionType }) => box => spObjects => (
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()}`),
	spObjects
)

const execute = box => cacheLeaf => actionType => spObjectGetter => (opts = {}) => box
	.chainAsync(async element => {
		let needSubsites;
		const { cached } = opts;
		const url = element.Url;
		let webUrl = url;
		const contextUrlSplits = pipe([popSlash, urlSplit])(url);
		if (actionType === 'create') {
			webUrl = getParentUrl(url);
			needSubsites = true;
		} else {
			needSubsites = isStringEmpty(url) || hasUrlTailSlash(url);
		}
		const clientContext = getClientContext(webUrl);
		const spObject = spObjectGetter({
			spContextObject: needSubsites ? getSPObjectCollection(clientContext) : getSPObject(clientContext),
			element: element
		});
		const cachePath = [...contextUrlSplits, NAME, needSubsites ? (cacheLeaf + 'Collection') : cacheLeaf];
		ACTION_TYPES_TO_UNSET[actionType] && getParentUrl, urlSplit, cache.unset(slice(0, -3)(cachePath));
		if (actionType === 'delete' || actionType === 'recycle') {
			return executorJSOM(clientContext)(opts);
		} else {
			const spObjectCached = cached ? cache.get(cachePath) : null;
			if (cached && spObjectCached) {
				return spObjectCached;
			} else {
				const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts);
				cache.set(currentSPObjects)(cachePath)
				return currentSPObjects;
			}
		}
	})
	.then(report({ ...opts, actionType })(box))
	.then(prepareResponseJSOM(opts))


// Interface

export default urls => {
	const instance = {
		box: getInstance(ContextUrlBox)(urls),
		getSPObject,
		getSPObjectCollection
	};
	const executeBinded = execute(instance.box)('properties');
	return {
		recycleBin: recycleBin(instance),
		// search: (instance => elements => search(instance, elements))(instance),
		list: (instance => elements => list(instance, elements))(instance),
		folder: (instance => elements => folder(instance, elements))(instance),
		file: (instance => elements => file(instance, elements))(instance),
		get: executeBinded(null)(prop('spContextObject')),
		create: executeBinded('create')(({ spContextObject, element }) => pipe([
			getInstanceEmpty,
			setFields({
				set_title: element.Title || void 0,
				set_description: element.Description,
				set_language: 1033,
				set_url: element.Title || void 0,
				set_useSamePermissionsAsParentSite: true,
				set_webTemplate: element.WebTemplate || APP_WEB_TEMPLATE
			}),
			methodI('add')(spContextObject),
			overstep(setFields({
				set_alternateCssUrl: element.AlternateCssUrl,
				set_associatedMemberGroup: element.AssociatedMemberGroup,
				set_associatedOwnerGroup: element.AssociatedOwnerGroup,
				set_associatedVisitorGroup: element.AssociatedVisitorGroup,
				set_customMasterUrl: element.CustomMasterUrl,
				set_enableMinimalDownload: element.EnableMinimalDownload,
				set_masterUrl: element.MasterUrl,
				set_objectVersion: element.ObjectVersion,
				set_quickLaunchEnabled: element.QuickLaunchEnabled,
				set_saveSiteAsTemplateEnabled: element.SaveSiteAsTemplateEnabled,
				set_serverRelativeUrl: element.ServerRelativeUrl,
				set_siteLogoUrl: element.SiteLogoUrl,
				set_syndicationEnabled: element.SyndicationEnabled,
				set_treeViewEnabled: element.TreeViewEnabled,
				set_uiVersion: element.UiVersion,
				set_uiVersionConfigurationEnabled: element.UiVersionConfigurationEnabled
			})),
			overstep(methodEmpty('update')),
		])(SP.WebCreationInformation)),
		update: executeBinded('update')(({ spContextObject, element }) => pipe([
			setFields({
				set_title: element.Title || void 0,
				set_description: element.Description,
				set_alternateCssUrl: element.AlternateCssUrl,
				set_associatedMemberGroup: element.AssociatedMemberGroup,
				set_associatedOwnerGroup: element.AssociatedOwnerGroup,
				set_associatedVisitorGroup: element.AssociatedVisitorGroup,
				set_customMasterUrl: element.CustomMasterUrl,
				set_enableMinimalDownload: element.EnableMinimalDownload,
				set_masterUrl: element.MasterUrl,
				set_objectVersion: element.ObjectVersion,
				set_quickLaunchEnabled: element.QuickLaunchEnabled,
				set_saveSiteAsTemplateEnabled: element.SaveSiteAsTemplateEnabled,
				set_serverRelativeUrl: element.ServerRelativeUrl,
				set_siteLogoUrl: element.SiteLogoUrl,
				set_syndicationEnabled: element.SyndicationEnabled,
				set_treeViewEnabled: element.TreeViewEnabled,
				set_uiVersion: element.UiVersion,
				set_uiVersionConfigurationEnabled: element.UiVersionConfigurationEnabled
			}),
			overstep(methodEmpty('update')),
		])(spContextObject)
		),
		delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spContextObject }) => {
			try { spContextObject.deleteObject() } catch (err) { new Error('Context url is wrong') }
			return spContextObject;
		})(opts),
	}
}