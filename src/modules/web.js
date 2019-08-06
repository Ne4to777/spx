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
	rootReport,
	isStrictUrl,
	switchCase,
	typeOf,
	removeEmptyUrls,
	removeDuplicatedUrls,
	shiftSlash,
	getTitleFromUrl,
	mergeSlashes
} from '../lib/utility'


// import search from './../modules/search'
import list from './list'
import folder from './folderWeb'
import file from './fileWeb'
import recycleBin from './recycleBin'
import user from './user'
import tag from './tag'
import mail from './mail'
import time from './time'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const lifter = switchCase(typeOf)({
	object: context => {
		const newContext = Object.assign({}, context)
		if (!context.Url && context.Title) newContext.Url = context.Title
		if (context.Url !== '/') newContext.Url = shiftSlash(newContext.Url)
		if (!context.Title && context.Url) newContext.Title = getTitleFromUrl(context.Url)
		return newContext
	},
	string: (contextUrl = '') => ({
		Url: contextUrl === '/' ? '/' : shiftSlash(mergeSlashes(contextUrl)),
		Title: getTitleFromUrl(contextUrl)
	}),
	default: () => ({
		Url: '',
		Title: ''
	})
})

class Web {
	constructor(urls) {
		this.name = 'web'
		this.box = getInstance(AbstractBox)(urls, lifter, arrayValidator)
		this.id = MD5(new Date().getTime()).toString()
	}

	async	get(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = element.Url
			const clientContext = getClientContext(elementUrl)
			const spObject = hasUrlTailSlash(elementUrl)
				? this.getSPObjectCollection(clientContext)
				: this.getSPObject(clientContext)
			return executeJSOM(clientContext, spObject, opts)
		})
		return prepareResponseJSOM(result, opts)
	}

	async create(opts) {
		const result = await this.box.chain(element => {
			const parentElementUrl = getParentUrl(element.Url)
			const clientContext = getClientContext(parentElementUrl)
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

			return executeJSOM(clientContext, spObject, opts)
		})

		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
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
			return executeJSOM(clientContext, spObject, opts)
		})
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
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
			await executorJSOM(clientContext, opts)
			return elementUrl
		})
		this.report('delete', opts)
		return prepareResponseJSOM(result, opts)
	}

	async	doesUserHavePermissions(type = 'fullMask') {
		const result = await this.box.chain(async element => {
			const clientContext = getClientContext(element.Url)
			const ob = getInstanceEmpty(SP.BasePermissions)
			ob.set(SP.PermissionKind[type])
			const spObject = this.getSPObject(clientContext).doesUserHavePermissions(ob)
			await executorJSOM(clientContext)
			return spObject.get_value()
		})
		return result
	}

	async	getSite(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const currentSPObjects = await executeJSOM(clientContext, spObject, opts)
		return prepareResponseJSOM(currentSPObjects, opts)
	}

	async	getCustomListTemplates(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const templates = spObject.getCustomListTemplates(clientContext.get_web())
		const currentSPObjects = await executeJSOM(clientContext, templates, opts)
		return prepareResponseJSOM(currentSPObjects, opts)
	}

	async	getWebTemplates(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const templates = spObject.getWebTemplates(1033, 0)
		const currentSPObjects = await executeJSOM(clientContext, templates, opts)
		return prepareResponseJSOM(currentSPObjects, opts)
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
	// 	return search(this, elements)
	// }

	tag(elements) {
		return tag(this, elements)
	}

	user(elements) {
		return user(this, elements)
	}

	mail(elements) {
		return mail(this, elements)
	}

	time() {
		return time(this)
	}

	getSPObject(clientContext) {
		return clientContext.get_web()
	}

	getSiteSPObject(clientContext) {
		return clientContext.get_site()
	}

	getSPObjectCollection(clientContext) {
		return this.getSPObject(clientContext).get_webs()
	}

	report(actionType, opts = {}) {
		rootReport(actionType, {
			...opts,
			name: this.name,
			box: this.box
		})
	}

	of(urls) {
		return getInstance(this.constructor)(urls)
	}
}

export default getInstance(Web)
