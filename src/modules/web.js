/* eslint class-methods-use-this:0 */
import { MD5 } from 'crypto-js'
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
	contextReport,
	isStrictUrl
} from '../lib/utility'


// import search from './../modules/search'
import list from './list'
import folder from './folderWeb'
import file from './fileWeb'
import recycleBin from './recycleBin'
import user from './user'
import tag from './tag'
import email from './email'
import * as time from './time'

class Web {
	constructor(urls) {
		this.box = getInstance(AbstractBox)(urls)
		this.id = MD5(new Date().getTime()).toString()
	}

	async	get(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = element.Url
			const clientContext = getClientContext(elementUrl)
			const spObject = hasUrlTailSlash(elementUrl)
				? this.getSPObjectCollection(clientContext)
				: this.getSPObject(clientContext)
			return executeJSOM(clientContext)(spObject)(opts)
		})
		return prepareResponseJSOM(opts)(result)
	}

	async create(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = getParentUrl(element.Url)
			if (!isStrictUrl(elementUrl)) return undefined
			const clientContext = getClientContext(elementUrl)

			const spObject = pipe([
				getInstanceEmpty,
				setFields({
					set_title: element.Title || undefined,
					set_description: element.Description,
					set_language: 1033,
					set_url: element.Title || undefined,
					set_useSamePermissionsAsParentSite: true,
					set_webTemplate: element.WebTemplate
				}),
				methodI('add')(this.getSPObjectCollection(clientContext)),
				overstep(
					setFields({
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
				),
				overstep(methodEmpty('update'))
			])(SP.WebCreationInformation)

			return executeJSOM(clientContext)(spObject)(opts)
		})
		this.report('create', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async update(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = element.Url
			if (!isStrictUrl(elementUrl)) return undefined
			const clientContext = getClientContext(elementUrl)
			const spObject = pipe([
				setFields({
					set_title: element.Title || undefined,
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
				overstep(methodEmpty('update'))
			])(this.getSPObject(clientContext))
			return executeJSOM(clientContext)(spObject)(opts)
		})
		this.report('update', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async delete(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = element.Url
			if (!isStrictUrl(elementUrl)) return undefined
			const clientContext = getClientContext(elementUrl)
			const spObject = this.getSPObject(clientContext)
			try {
				spObject.deleteObject()
			} catch (err) {
				return new Error('Context url is wrong')
			}
			await executorJSOM(clientContext)(opts)
			return elementUrl
		})
		this.report('delete', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	doesUserHavePermissions(type = 'fullMask') {
		const result = await this.box.chain(async element => {
			const clientContext = getClientContext(element.Url)
			const ob = getInstanceEmpty(SP.BasePermissions)
			ob.set(SP.PermissionKind[type])
			const spObject = this.getSPObject(clientContext).doesUserHavePermissions(ob)
			await executorJSOM(clientContext)()
			return spObject.get_value()
		})
		return result
	}

	getSPObject(clientContext) {
		return methodEmpty('get_web')(clientContext)
	}

	getSiteSPObject(clientContext) {
		return methodEmpty('get_site')(clientContext)
	}

	getSPObjectCollection(clientContext) {
		return pipe([this.getSPObject, methodEmpty('get_webs')])(clientContext)
	}


	report(actionType, opts = {}) {
		contextReport(actionType, { ...opts, NAME: 'web', box: this.box })
	}

	async	getSite(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const currentSPObjects = await executeJSOM(clientContext)(spObject)(opts)
		return prepareResponseJSOM(opts)(currentSPObjects)
	}

	async	getCustomListTemplates(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const templates = spObject.getCustomListTemplates(clientContext.get_web())
		const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
		return prepareResponseJSOM(opts)(currentSPObjects)
	}

	async	getWebTemplates(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const templates = spObject.getWebTemplates(1033, 0)
		const currentSPObjects = await executeJSOM(clientContext)(templates)(opts)
		return prepareResponseJSOM(opts)(currentSPObjects)
	}

	recycleBin() {
		return recycleBin(this)
	}

	list(elements) {
		return list(this, elements, 'list')
	}

	library(elements) {
		return list(this, elements, 'library')
	}

	folder(elements) {
		return folder(this, elements)
	}

	file(elements) {
		return file(this, elements)
	}

	// search(elements) {
	// 	return search(this,elements)
	// }

	tag(elements) {
		return tag(this, elements)
	}

	user(elements) {
		return user(this, elements)
	}

	email(elements) {
		return email(this, elements)
	}

	time() {
		return time(this)
	}
}

export default getInstance(Web)
