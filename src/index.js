/* eslint max-len:0 */
/* eslint no-unused-vars:0 */
/* eslint no-restricted-syntax:0 */
/* eslint import/no-extraneous-dependencies:0 */
import axios from 'axios'
import $ from 'jquery'
import test from './test/index'

import spx from './modules/web'
import {
	getCamlQuery, getCamlView, camlLog, craftQuery, concatQueries
} from './lib/query-parser'
import {
	log,
	executeJSOM,
	prepareResponseJSOM,
	getClientContext,
	executorJSOM,
} from './lib/utility'
import * as cache from './lib/cache'
// import privateData from '../dev/private.json'

// spx.setCustomUsersList({
// 	webTitle: privateData.customUsersWeb,
// 	listTitle: privateData.customUsersList
// })

window.axios = axios
window.log = log
window.getCamlView = getCamlView
window.getCamlQuery = getCamlQuery
window.craftQuery = craftQuery
window.concatQueries = concatQueries
window.camlLog = camlLog
window.spx = spx

window.cache = cache

spx().user().setDefaults({
	customWebTitle: 'AM',
	customListTitle: 'UsersAD',
	customIdColumn: 'uid',
	customLoginColumn: 'Login',
	customNameColumn: 'Title',
	customEmailColumn: 'Email',
	customQuery: 'Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))'
})

test()

$('#send').click(e => {
	e.preventDefault()
	console.log($('#file').get(0).files[0])
	spx('/test/spx')
		.library('Files')
		.file({
			Content: $('#file').get(0).files[0],
			OnProgress: console.log,
			Folder: 'a'
		})
		.create()
		.then(log)
})

const getWebPermission = () => {
	const clientContext = new SP.ClientContext('/common')
	const web = clientContext.get_web()
	const ob = new SP.BasePermissions()
	ob.set(SP.PermissionKind.fullMask)
	const per = web.doesUserHavePermissions(ob)
	clientContext.executeQueryAsync(() => console.log(per.get_value()))
}

// getWebPermission()

const perms = {
	emptyMask: 0,
	viewListItems: 1,
	addListItems: 2,
	editListItems: 3,
	deleteListItems: 4,
	approveItems: 5,
	openItems: 6,
	viewVersions: 7,
	deleteVersions: 8,
	cancelCheckout: 9,
	managePersonalViews: 10,
	manageLists: 12,
	viewFormPages: 13,
	anonymousSearchAccessList: 14,
	open: 17,
	viewPages: 18,
	addAndCustomizePages: 19,
	applyThemeAndBorder: 20,
	applyStyleSheets: 21,
	viewUsageData: 22,
	createSSCSite: 23,
	manageSubwebs: 24,
	createGroups: 25,
	managePermissions: 26,
	browseDirectories: 27,
	browseUserInfo: 28,
	addDelPrivateWebParts: 29,
	updatePersonalWebParts: 30,
	manageWeb: 31,
	anonymousSearchAccessWebLists: 32,
	useClientIntegration: 37,
	useRemoteAPIs: 38,
	manageAlerts: 39,
	createAlerts: 40,
	editMyUserInfo: 41,
	enumeratePermissions: 63,
	fullMask: 65
}

const getListPermission = () => {
	const clientContext = new SP.ClientContext('/common')
	const web = clientContext.get_web()
	const oList = web.get_lists().getByTitle('Administrators')
	clientContext.load(oList, 'EffectiveBasePermissions')
	clientContext.executeQueryAsync(() => {
		console.log(oList.get_effectiveBasePermissions().has(SP.PermissionKind.manageWeb))
	}, log)
}

// getListPermission();

const removeUserFromGroup = async () => {
	const ctx = getClientContext('/')
	const web = ctx.get_web()

	const group = web.get_siteGroups().getByName('Everyone')
	// group.get_users().removeByLoginName(userLoginName);
	await executeJSOM(ctx)(group)()
	console.log(group)
}

// removeUserFromGroup()

const retrieveAllUsersInGroup = async () => {
	const clientContext = getClientContext('/')
	console.log(clientContext.get_web())
	const collGroup = clientContext.get_web().get_siteGroups()
	const oGroup = collGroup.getByName('pps_administrators')
	const collUser = oGroup.get_users()
	clientContext.load(collUser)
	await executorJSOM(clientContext)()
	console.log(collUser)
	let userInfo = ''

	const userEnumerator = collUser.getEnumerator()
	while (userEnumerator.moveNext()) {
		const oUser = userEnumerator.get_current()
		userInfo += `\nUser: ${oUser.get_title()}\nID: ${oUser.get_id()}\nEmail: ${oUser.get_email()}\nLogin Name: ${oUser.get_loginName()}`
	}
	console.log(userInfo)
}

// retrieveAllUsersInGroup()

const getAllGroups = async () => {
	const clientContext = getClientContext('/')
	const siteGroups = clientContext.get_web().get_siteGroups()
	const groups = await executeJSOM(clientContext)(siteGroups)()
	return prepareResponseJSOM()(groups)
}

const getGroupById = async id => {
	const clientContext = getClientContext('/')
	const siteGroups = clientContext.get_web().get_siteGroups()
	const group = siteGroups.getById(id)
	await executeJSOM(clientContext)(group)()
	return prepareResponseJSOM()(group)
}
// getGroupById(36).then(log);
// getAllGroups().then(log)

const getGroupOwnerById = async id => {
	const clientContext = getClientContext('/')
	const siteGroups = clientContext.get_web().get_siteGroups()
	const group = siteGroups.getById(id)
	const owner = group.get_owner()
	await executorJSOM(clientContext)()
	console.log(owner)
	return prepareResponseJSOM()(owner)
}

// getGroupOwnerById(36).then(log)

const setGroupOwnerById = async id => {
	const clientContext = getClientContext('/')
	const siteGroups = clientContext.get_web().get_siteGroups()
	const group = siteGroups.getById(id)
	const owner = group.get_owner()
	group.set_owner(owner)
	executorJSOM(clientContext)()
}

// setGroupOwnerById(36).then(log)

const removeGroupById = async id => {
	const clientContext = getClientContext('/')
	const siteGroups = clientContext.get_web().get_siteGroups()
	siteGroups.removeById(id)
	await executorJSOM(clientContext)()
}

// removeGroupById(16220)

const removeGroupsByIds = async ids => {
	for (const id of ids) {
		await removeGroupById(id)
		// console.log(id);
	}
	console.log('done')
}

const removeGroups = async () => {
	const groupsToExclude = {
		Administrators: true,
		Developers: true,
		Everyone: true
	}
	const groups = await getAllGroups()
	console.log(groups)
	const filtereds = groups.filter(el => !groupsToExclude[el.LoginName])
	console.log(filtereds)
	const ids = filtereds.map(el => el.Id)
	console.log(ids)
	// removeGroupsByIds(ids)
}
// removeGroups()

// const APP_WEB_TEMPLATE = '{E2A30D74-39CB-429E-A5E0-4C775BE848CE}#Default'

// spx('test/spx').list('Keywords').column('TaxKeyword').get().then(log)
// spx('test/spx').list('Keywords').item(2).get().then(item => {
// console.log(item);
// 'TaxKeywordTaxHTField': 'a|b1bcd272-e0e4-47b0-9560-a63c0be091ad;b2b|34e20559-f098-48f1-b9af-93ac54b0841b'
// spx('test/spx').list('Keywords').item({ 'TaxKeyword': item.TaxKeyword }).create().then(log)
// spx('test/spx').list('Keywords').item({ 'TaxKeyword': ['a|b1bcd272-e0e4-47b0-9560-a63c0be091ad', 'b2b|34e20559-f098-48f1-b9af-93ac54b0841b'] }).create().then(log)
// spx('test/spx').list('Keywords').item({ 'TaxKeyword': 'a|b1bcd272-e0e4-47b0-9560-a63c0be091ad;b2b|34e20559-f098-48f1-b9af-93ac54b0841b' }).create().then(log)
// })

// spx.mail.send({ from: 'AlekseyAlekseew@dme.ru', to: ['AlekseyAlekseew@dme.ru'], subject: 'Test Subject', body: 'Test body' }).then(console.log)

// spx.time.get().then(log)
