import * as utility from './../utility';
import * as cache from './../cache';

import List from './../modules/list';
import Folder from './../modules/folderWeb';
import File from './../modules/fileWeb';
import User from './../modules/user';
import RecycleBin from './../modules/recycleBin';

export default class Web {
	constructor(contextUrl) {
		this._contextUrlIsArray = typeOf(contextUrl) === 'array';
		this._contextUrls = this._contextUrlIsArray ? contextUrl : [contextUrl || '/'];
	}

	get _APP_WEB_TEMPLATE() { return '{FF0F2C5D-45AB-4B50-8ABC-883489A45C44}#app' }

	// Interface

	folder(elementUrl) { return new Folder(this, elementUrl) }
	file(elementUrl) { return new File(this, elementUrl) }
	user(elementUrl) { return new User(this, elementUrl) }
	list(elementUrl) { return new List(this, elementUrl) }
	get recycleBin() { return new RecycleBin(this) }

	async get(opts) {
		return this._execute(null, spContextObject =>
			(spContextObject.cachePath = spContextObject.getEnumerator ? 'propertiesCollection' : 'properties', spContextObject), opts)
	}

	async getTemplates(opts) {
		return this._execute(null, spContextObject => {
			const spObject = spContextObject.getAvailableWebTemplates(1033, false);
			spObject.cachePath = 'templates';
			return spObject;
		}, opts);
	}

	async getListTemplates(opts) {
		return this._execute(null, spContextObject => {
			const spObject = spContextObject.get_listTemplates();
			spObject.cachePath = 'listTemplates';
			return spObject;
		}, opts);
	}

	async create(opts = {}) {
		return this._execute('create', (spContextObject, element) => {
			const webCreationInfo = new SP.WebCreationInformation;
			utility.setFields(webCreationInfo, {
				set_title: element.Title,
				set_description: element.Description,
				set_language: 1033,
				set_url: element.Title,
				set_useSamePermissionsAsParentSite: true,
				set_webTemplate: element.WebTemplate || this._APP_WEB_TEMPLATE
			})
			const spObject = spContextObject.add(webCreationInfo);
			spObject.cachePath = 'properties';
			return spObject;
		}, opts);
	}
	async update(opts) {
		return this._execute('update', (spContextObject, element) => {
			utility.setFields(spContextObject, {
				set_title: element.Title,
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
			})
			spContextObject.update();
			spContextObject.cachePath = 'properties';
			return spContextObject;
		}, opts);
	}
	async delete(opts) {
		return this._execute('delete', spContextObject => {
			try { spContextObject.deleteObject() } catch (err) { throw new Error('Context url is wrong') }
			spContextObject.cachePath = 'properties';
			return spContextObject;
		}, opts);
	}

	// Internal

	get _name() { return 'web' }

	async _execute(actionType, spObjectGetter, opts = {}) {
		const { cached } = opts;
		const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
			const context = this._liftElementUrlType(contextUrl);
			const contextUrlSplits = context.Url.split('/');
			let needSubsites = contextUrlSplits.slice(-1) === '/';
			let webUrl = context.Url;
			if (actionType === 'create') {
				webUrl = contextUrlSplits.slice(0, -1).join('/') || '/';
				needSubsites = true;
			}
			const clientContext = utility.getClientContext(webUrl);
			const spObject = spObjectGetter(this._getSPObject(clientContext, needSubsites), context);
			const cachePaths = [...contextUrlSplits, this._name, spObject.cachePath];
			utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset(contextUrlSplits.slice(0, -1));
			if (actionType === 'delete' || actionType === 'recycle') {
				await utility.executeQueryAsync(clientContext, opts);
			} else {
				const spObjectCached = cached ? cache.get(cachePaths) : null;
				if (cached && spObjectCached) {
					return spObjectCached
				} else {
					const spObjects = utility.load(clientContext, spObject, opts);
					await utility.executeQueryAsync(clientContext, opts);
					cache.set(cachePaths, spObjects);
					return spObjects
				}
			}
		}))
		this._log(actionType, opts);
		opts.isArray = this._contextUrlIsArray;
		return utility.prepareResponseJSOM(elements, opts);
	}

	_getSPObject(clientContext, needSubsites) {
		if (clientContext) {
			const web = clientContext.get_web();
			return needSubsites ? web.get_webs() : web;
		} else {
			return clientContext.get_site().get_rootWeb().get_webs();
		}
	}

	_liftElementUrlType(contextUrl) {
		switch (typeOf(contextUrl)) {
			case 'object':
				if (!contextUrl.Url) contextUrl.Url = contextUrl.Title;
				if (!contextUrl.Title) contextUrl.Title = utility.getLastPath(contextUrl.Url);
				return contextUrl;
			case 'string':
				contextUrl = contextUrl.replace(/\/$/, '');
				return {
					Url: contextUrl,
					Title: utility.getLastPath(contextUrl)
				}
		}
	}

	_log(actionType, opts = {}) {
		!opts.silent && actionType &&
			console.log(`${
				utility.ACTION_TYPES[actionType]} ${
				this._name}: ${
				this._contextUrls.map(el => this._liftElementUrlType(el).Title).join(', ')}`);
	}
}