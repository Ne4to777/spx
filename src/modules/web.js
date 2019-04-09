import {
	AbstractBox,
	getClientContext,
	prepareResponseJSOM,
	setFields,
	pipe,
	methodEmpty,
	getInstance,
	overstep,
	executeJSOM,
	executorJSOM,
	getInstanceEmpty,
	methodI,
	getParentUrl,
	hasUrlTailSlash,
	switchCase,
	typeOf,
	mergeSlashes,
	ifThen,
	constant,
	shiftSlash,
	isArrayFilled,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	getTitleFromUrl,
	contextReport,
	isStrictUrl
} from './../lib/utility'

import {
	MD5
} from 'crypto-js';

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

const liftWebType = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context);
		if (!context.Url && context.Title) newContext.Url = context.Title;
		if (context.Url !== '/') newContext.Url = shiftSlash(newContext.Url);
		if (!context.Title && context.Url) newContext.Title = getTitleFromUrl(context.Url);
		return newContext
	},
	string: (contextUrl = '') => ({
		Url: contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl)),
		Title: getTitleFromUrl(contextUrl)
	}),
	default: _ => ({
		Url: '',
		Title: ''
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value);
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([
					map(liftWebType),
					removeEmptyUrls,
					removeDuplicatedUrls
				]),
				constant([liftWebType()])
			])(value)
			: liftWebType(value);
	}
}

// Interface

export default urls => {
	const instance = {
		box: getInstance(Box)(urls),
		getSPObject,
		getSPObjectCollection,
		id: MD5(new Date().getTime()).toString()
	};
	const report = actionType => (opts = {}) => contextReport({ ...opts, NAME, actionType, box: instance.box });
	return {
		recycleBin: recycleBin(instance),
		// search: search(instance),
		list: list('list')(instance),
		library: list('library')(instance),
		folder: folder(instance),
		file: file(instance),


		get: async opts => {
			const result = await instance.box.chain(async element => {
				const elementUrl = element.Url;
				const clientContext = getClientContext(elementUrl);
				const spObject = hasUrlTailSlash(elementUrl) ? getSPObjectCollection(clientContext) : getSPObject(clientContext);
				return executeJSOM(clientContext)(spObject)(opts);
			})
			return prepareResponseJSOM(opts)(result)
		},

		create: async opts => {
			const result = await instance.box.chain(async element => {
				const elementUrl = getParentUrl(element.Url);
				if (!isStrictUrl(elementUrl)) return;
				const clientContext = getClientContext(elementUrl);

				const spObject = pipe([
					getInstanceEmpty,
					setFields({
						set_title: element.Title || void 0,
						set_description: element.Description,
						set_language: 1033,
						set_url: element.Title || void 0,
						set_useSamePermissionsAsParentSite: true,
						set_webTemplate: element.WebTemplate || APP_WEB_TEMPLATE
					}),
					methodI('add')(getSPObjectCollection(clientContext)),
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
				])(SP.WebCreationInformation)

				return executeJSOM(clientContext)(spObject)(opts);
			})
			report('create')(opts);
			return prepareResponseJSOM(opts)(result)
		},

		update: async opts => {
			const result = await instance.box.chain(async element => {
				const elementUrl = element.Url;
				if (!isStrictUrl(elementUrl)) return;
				const clientContext = getClientContext(elementUrl);
				const spObject = pipe([
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
				])(getSPObject(clientContext))
				return await executeJSOM(clientContext)(spObject)(opts);
			})
			report('update')(opts);
			return prepareResponseJSOM(opts)(result)
		},

		delete: async opts => {
			const result = await instance.box.chain(async element => {
				const elementUrl = element.Url;
				if (!isStrictUrl(elementUrl)) return;
				const clientContext = getClientContext(elementUrl);
				const spObject = getSPObject(clientContext);
				try { spObject.deleteObject() } catch (err) { new Error('Context url is wrong') }
				await executorJSOM(clientContext)(opts);
				return elementUrl
			})
			report('delete')(opts);
			return prepareResponseJSOM(opts)(result)
		},

		doesUserHavePermissions: async (type = 'fullMask') => {
			const result = await instance.box.chain(async element => {
				const clientContext = getClientContext(element.Url);
				const ob = getInstanceEmpty(SP.BasePermissions);
				ob.set(SP.PermissionKind[type]);
				const spObject = getSPObject(clientContext).doesUserHavePermissions(ob);
				await executorJSOM(clientContext)();
				return spObject.get_value()
			})
			return result;
		}
	}
}