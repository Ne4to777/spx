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
	switchType,
	removeEmptyUrls,
	removeDuplicatedUrls,
	shiftSlash,
	getTitleFromUrl,
	getPermissionMasks
} from '../lib/utility'


import search from './search'
import list from './list'
import folder from './folderWeb'
import file from './fileWeb'
import recycleBin from './recycleBin'
import user from './user'
import group from './group'
import keyword from './keyword'
import mail from './mail'
import time from './time'

const KEY_PROP = 'Url'

const arrayValidator = pipe([removeEmptyUrls, removeDuplicatedUrls])

const lifter = switchType({
	object: context => {
		const newContext = Object.assign({ [KEY_PROP]: '', Title: '' }, context)
		if (!context[KEY_PROP] && context.Title) newContext[KEY_PROP] = context.Title
		if (context[KEY_PROP] !== '/') newContext[KEY_PROP] = shiftSlash(newContext[KEY_PROP])
		if (!context.Title && context[KEY_PROP]) newContext.Title = getTitleFromUrl(context[KEY_PROP])
		newContext[KEY_PROP] = newContext[KEY_PROP].replace('file:///', '')
		return newContext
	},
	string: (contextUrl = '') => ({
		[KEY_PROP]: contextUrl === '/' ? '/' : shiftSlash(contextUrl.replace('file:///', '')),
		Title: getTitleFromUrl(contextUrl)
	}),
	default: () => ({
		[KEY_PROP]: '',
		Title: ''
	})
})

class Web {
	constructor(urls) {
		this.name = 'web'
		this.box = getInstance(AbstractBox)(urls, lifter, arrayValidator)
		this.id = MD5(new Date().getTime()).toString()
	}

	async get(opts) {
		const result = await this.box.chain(async element => {
			const elementUrl = element[KEY_PROP]
			const clientContext = getClientContext(elementUrl)
			const spObject = hasUrlTailSlash(elementUrl)
				? this.getSPObjectCollection(clientContext)
				: this.getSPObject(clientContext)
			return executeJSOM(clientContext, spObject, opts)
		})
		return prepareResponseJSOM(result, opts)
	}

	async create(opts) {
		const values = this.box.getIterable()
		const result = []
		for (let i = 0; i < values.length; i += 1) {
			const element = values[i]
			const parentElementUrl = getParentUrl(element[KEY_PROP])
			const clientContext = getClientContext(parentElementUrl)
			const spObject = pipe([
				getInstanceEmpty,
				setFields({
					set_title: element.Title,
					set_description: element.Description,
					set_language: element.Language || 1033,
					set_url: element.Title,
					set_useSamePermissionsAsParentSite: element.UseSamePermissionsAsParentSite || true,
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

			try {
				result.push(await executeJSOM(clientContext, spObject, opts))
			} catch (err) {
				throw err
			}
		}

		this.report('create', opts)
		return prepareResponseJSOM(this.box.isArray() ? result : result[0], opts)
	}

	async update(opts) {
		const values = this.box.getIterable()
		const result = []
		for (let i = 0; i < values.length; i += 1) {
			const element = values[i]
			const elementUrl = element[KEY_PROP]
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
			try {
				result.push(await executeJSOM(clientContext, spObject, opts))
			} catch (err) {
				throw err
			}
		}
		this.report('update', opts)
		return prepareResponseJSOM(this.box.isArray() ? result : result[0], opts)
	}

	async delete(opts) {
		const values = this.box.getIterable()
		for (let i = 0; i < values.length; i += 1) {
			const element = values[i]
			const elementUrl = element[KEY_PROP]
			if (!isStrictUrl(elementUrl)) throw new Error('Wrong context url')
			const clientContext = getClientContext(elementUrl)
			const spObject = this.getSPObject(clientContext)
			try {
				spObject.deleteObject()
				await executorJSOM(clientContext)
			} catch (err) {
				throw err
			}
		}
		this.report('delete', opts)
		return undefined
	}

	async doesUserHavePermissions(type = 'fullMask') {
		const result = await this.box.chain(async element => {
			const clientContext = getClientContext(element[KEY_PROP])
			const ob = getInstanceEmpty(SP.BasePermissions)
			const permissionId = getPermissionMasks()[type]
			if (permissionId === undefined) throw new Error('Permission mask has invalid value')
			ob.set(permissionId)
			const spObject = this.getSPObject(clientContext).doesUserHavePermissions(ob)
			await executorJSOM(clientContext)
			return spObject.get_value()
		})
		return result
	}

	async getPermissions() {
		const result = await this.box.chain(async element => {
			const clientContext = getClientContext(element[KEY_PROP])
			return executeJSOM(clientContext, this.getSPObject(clientContext), { view: 'EffectiveBasePermissions' })
		})

		const allPermissionMasks = getPermissionMasks()
		return Object.keys(allPermissionMasks).reduce((acc, el) => {
			acc[el] = result.get_effectiveBasePermissions().has(allPermissionMasks[el])
			return acc
		}, {})
	}

	async breakRoleInheritance() {
		return this.box.chain(async element => {
			const clientContext = getClientContext(element[KEY_PROP])
			const spObject = this.getSPObject(clientContext)
			methodEmpty('breakRoleInheritance')(spObject)
			return executorJSOM(clientContext)
		})
	}

	async resetRoleInheritance() {
		return this.box.chain(async element => {
			const clientContext = getClientContext(element[KEY_PROP])
			const spObject = this.getSPObject(clientContext)
			methodEmpty('resetRoleInheritance')(spObject)
			return executorJSOM(clientContext)
		})
	}

	async getSite(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const currentSPObjects = await executeJSOM(clientContext, spObject, opts)
		return prepareResponseJSOM(currentSPObjects, opts)
	}

	async getCustomListTemplates(opts) {
		const clientContext = getClientContext('/')
		const spObject = this.getSiteSPObject(clientContext)
		const templates = spObject.getCustomListTemplates(clientContext.get_web())
		const currentSPObjects = await executeJSOM(clientContext, templates, opts)
		return prepareResponseJSOM(currentSPObjects, opts)
	}

	async getWebTemplates(opts) {
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

	search(elements) {
		return search(this, elements)
	}

	keyword(elements) {
		return keyword(this, elements)
	}

	user(elements) {
		return user(this, elements)
	}

	group(elements) {
		return group(this, elements)
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
