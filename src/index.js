import test from './test/index.js'
import axios from 'axios';
import $ from 'jquery';

// import S from 'sanctuary'

// import * as utility from './utility';
// import {
// 	webBundle,
// 	folderBundle,
// 	fileBundle
// } from './test'
import spx from './modules/site'
import {
	getCamlQuery,
	getCamlView,
	camlLog,
	craftQuery,
	concatQueries,
} from './lib/query-parser';
import { log, executeJSOM, prepareResponseJSOM, prepareResponseREST, getClientContext, executorJSOM } from './lib/utility';
import * as cache from './lib/cache';
window.axios = axios;
window.log = log;
window.getCamlView = getCamlView;
window.getCamlQuery = getCamlQuery;
window.craftQuery = craftQuery;
window.concatQueries = concatQueries;
window.camlLog = camlLog;
window.spx = spx;


window.cache = cache;

test()

// axios({
// 	url: `/test/spx/_api/web/lists/getbytitle('Files')/rootfolder/folders/getbyurl('a')/files/getbyurl('add.png')`,
// 	headers: {
// 		'accept': 'application/json;odata=verbose',
// 		'content-type': 'application/json;odata=verbose'
// 	}
// }).then(res => {
// 	console.log(res.data.d);
// })


// const uid = Math.round(Math.random() * 10000);
// log(uid);

// const renderData_ = uid => {
// 	spx.user(uid).get().then(user => {
// 		log(user);
// 		spx('/Alekseev').list('Test').item(`Title Eq ${user.Gender === 'М' ? 'avay' : 'avya'}`).get().then(items => {
// 			log(items)
// 			$('body').html(`<div><div>${user.Title}</div><div>${items[0].Title}</div></div>`);
// 		})
// 	})
// }


// const renderData = uid => {
// 	$('body').html('');
// 	spx.user(uid).get()
// 		.then(user => ($('body').append(`<div>${user.Title}</div>`), user))
// 		.then(log)
// 		.then(user => `Title Eq ${user.Gender === 'М' ? 'avay' : 'avya'}`)
// 		.then(log)
// 		.then(spx('/Alekseev').list('Test').item)
// 		.then(log)
// 		.then(S.prop('get'))
// 		.then(log)
// 		.then(async fn => await fn())
// 		.then(log)
// 		.then(items => $('body').append(`<div>${items[0].Title}</div>`))
// }

// renderData(uid);

// let web = spx(['Alekseev', 'Intellect']);
// let site = spx.getCustomListTemplates().then(log)
// console.log(web);
// web.get().then(log)
// spx().recycleBin.get().then(log)
// console.log(cache.set(1)(['a', 'b']));
// console.log(cache.get(['a', '', 'b']));
// console.log(cache.unset(['a', '', 'b']));


// let clientContext = new SP.ClientContext('/')
// let site = clientContext.get_site();
// clientContext.load(site);
// clientContext.executeQueryAsync(_ => log(site))

// console.log(spx(['/a', '/b']));
// spx('Intellect/').get({ groupBy: 'Title' }).then(log)
// spx('test/spx').get().then(log)
// spx('test/spx1').get().then(log)


// spx('/test/spx').list('caml').get().then(log)

// function getCurrentUserPermission() {
// 	var web, clientContext, currentUser, oList, perMask;

// 	clientContext = new SP.ClientContext('/test/spx');
// 	web = clientContext.get_web();
// 	currentUser = web.get_currentUser();
// 	oList = web.get_lists().getByTitle('caml');
// 	clientContext.load(oList, 'EffectiveBasePermissions');
// 	clientContext.load(currentUser);
// 	clientContext.load(web);

// 	clientContext.executeQueryAsync(function () {
// 		if (oList.get_effectiveBasePermissions().has(SP.PermissionKind.editListItems)) {
// 			console.log("user has edit permission");
// 		} else {
// 			console.log("user doesn't have edit permission");
// 		}
// 	}, function (sender, args) {
// 		console.log('request failed ' + args.get_message() + '\n' + args.get_stackTrace());
// 	});
// }

// getCurrentUserPermission()

// spx('/test/spx').list('caml').column(['Title', 'Author']).get().then(log)

// spx('/test/spx').list('Test').item({ folder: 'a' }).get().then(log)

// spx('/test/spx').file('Files/a/test.txt').get({ asBlob: true }).then(log)
// spx('/test/spx').file('/test.txt').create().then(log)

$('#send').click(e => {
	e.preventDefault();
	console.log($('#file').get(0).files[0]);
	spx('/test/spx').library('Files').file({ Content: $('#file').get(0).files[0], OnProgress: console.log, Folder: 'a' }).create().then(log);
});

const getFile = _ => {
	const clientContext = new SP.ClientContext('/test/spx');
	const list = clientContext.get_web().get_lists().getByTitle('Files');
	const file = list.get_rootFolder().get_folders().getByUrl('a').get_files().getByUrl('test.txt');
	clientContext.load(file);
	clientContext.executeQueryAsync(_ => {
		console.log(file);

	}, log)
}

// getFile()

const asyncF = _ => new Promise(resolve => setTimeout(_ => resolve, 1000))

const iterAsync = async _ => {
	const array = [1, 2, 3]
	// for (const el of array) {
	// 	await asyncF()
	// 	console.log(el);
	// }
	array.map(async el => {
		await asyncF()
		console.log(el);
	})
}
// iterAsync()

const itemsToCreate = [];
for (let i = 0; i < 10; i++) {
	itemsToCreate.push({ Title: i });
}


// spx('test/spx').list('Test2').folder(['c', 'd']).create().then(log)
// spx('test/spx').list('Test').item(itemsToCreate).create().then(log)
// spx('test/spx').list('Test').item({ Query: '', Columns: { Title: null } }).updateByQuery().then(log)
// spx('test/spx').list('Test').item(['Title']).deleteEmpties().then(log)

// spx('test/spx').list('Test').item(['Title']).getDuplicates().then(log)
// spx('test/spx').list('Test').item(['Title']).deleteDuplicates().then(log)
// spx('test/spx').list('Test').item({ Query: 'ID Geq 128', Columns: ['Title'] }).erase().then(log)
// spx('test/spx').list('Test').item({ Key: 'ID', Forced: true, Target: { Column: 'Title1' }, Source: { Column: 'Title' }, Mediator: col => value => col + value }).merge({ expanded: true }).then(log)
// spx('test/spx').list('Posts').item({ ID: 4, Title: 'hi' }).createReply().then(log)
// spx('test/spx').list('Test').item({ Title: 'yvayva', Folder: 'a' }).create().then(log)

// spx('test/spx').list('Test').item([{
// 	Title: 'pvyay',
// 	Title1: 'pvp yvavy',
// 	lookup: [3, 4]
// }]).create().then(log)
// spx('test/spx').list('Test2').delete().then(log).catch(_ => _)

// new Promise(resolve => resolve(spx('test/spx').list))
// 	.then(list => list('Test').folder('avyayva').get())
// 	.then(log)

// pipe([methodEmpty('get'), method('then')(log)])(spx('test/spx').list('Test').item())

// spx('test/spx').list('Test').item().get().then(log)


// spx.tag('test').get().then(log)
// spx('Intellect').search('Search')().then(log)

// spx('test/spx/').get().then(webs => {
// 	log(webs)
// 	const webUrls = webs.map(web => web.ServerRelativeUrl);
// 	console.log(webUrls);
// 	spx(webUrls).delete()
// })

// spx('test/spx').get({ expanded: true }).then(web => {
// 	console.log(web);
// 	const props = web.get_allProperties();
// 	const clientContext = web.get_context();
// 	executeJSOM(clientContext)(props)()
// 		.then(log)
// 		.then(methodEmpty('get_objectData'))
// 		.then(log)
// 		.then(methodEmpty('get_methodReturnObjects'))
// 		.then(log)
// 		.then(prop('$m_dict'))
// 		.then(log)
// })


const getWebPermission = _ => {
	const clientContext = new SP.ClientContext('/common');
	const web = clientContext.get_web();
	const ob = new SP.BasePermissions();
	ob.set(SP.PermissionKind.fullMask);
	const per = web.doesUserHavePermissions(ob);
	clientContext.executeQueryAsync(_ => console.log(per.get_value()));
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


const getListPermission = _ => {

	const clientContext = new SP.ClientContext('/common');
	const web = clientContext.get_web();
	const oList = web.get_lists().getByTitle('Administrators');
	clientContext.load(oList, 'EffectiveBasePermissions');
	clientContext.executeQueryAsync(_ => {
		console.log(oList.get_effectiveBasePermissions().has(SP.PermissionKind.manageWeb));
	}, log);
}

// getListPermission();


const removeUserFromGroup = async _ => {
	const ctx = getClientContext('/');
	const web = ctx.get_web();

	const group = web.get_siteGroups().getByName('Everyone');
	// group.get_users().removeByLoginName(userLoginName);
	await executeJSOM(ctx)(group)()
	console.log(group);
}

// removeUserFromGroup()



const retrieveAllUsersInGroup = async _ => {

	const clientContext = getClientContext('/');
	console.log(clientContext.get_web());
	const collGroup = clientContext.get_web().get_siteGroups();
	const oGroup = collGroup.getByName('pps_administrators');
	const collUser = oGroup.get_users();
	clientContext.load(collUser);
	await executorJSOM(clientContext)()
	console.log(collUser);
	let userInfo = '';

	const userEnumerator = collUser.getEnumerator();
	while (userEnumerator.moveNext()) {
		const oUser = userEnumerator.get_current();
		userInfo += '\nUser: ' + oUser.get_title() +
			'\nID: ' + oUser.get_id() +
			'\nEmail: ' + oUser.get_email() +
			'\nLogin Name: ' + oUser.get_loginName();
	}
	console.log(userInfo);
}

// retrieveAllUsersInGroup()

const getAllGroups = async _ => {
	const clientContext = getClientContext('/');
	const siteGroups = clientContext.get_web().get_siteGroups();
	const groups = await executeJSOM(clientContext)(siteGroups)()
	return prepareResponseJSOM()(groups);
}

const getGroupById = async id => {
	const clientContext = getClientContext('/');
	const siteGroups = clientContext.get_web().get_siteGroups();
	const group = siteGroups.getById(id);
	await executeJSOM(clientContext)(group)();
	return prepareResponseJSOM()(group);
}
// getGroupById(36).then(log);
// getAllGroups().then(log)


const getGroupOwnerById = async id => {
	const clientContext = getClientContext('/');
	const siteGroups = clientContext.get_web().get_siteGroups();
	const group = siteGroups.getById(id);
	const owner = group.get_owner();
	await executorJSOM(clientContext)();
	console.log(owner);
	return prepareResponseJSOM()(owner);
}

// getGroupOwnerById(36).then(log)

const setGroupOwnerById = async id => {
	const clientContext = getClientContext('/');
	const siteGroups = clientContext.get_web().get_siteGroups();
	const group = siteGroups.getById(id);
	const owner = group.get_owner();
	group.set_owner(owner);
	executorJSOM(clientContext)();
}


// setGroupOwnerById(36).then(log)




const removeGroupById = async id => {
	const clientContext = getClientContext('/');
	const siteGroups = clientContext.get_web().get_siteGroups();
	siteGroups.removeById(id)
	await executorJSOM(clientContext)();
}

// removeGroupById(16220)

const removeGroupsByIds = async ids => {
	for (const id of ids) {
		await removeGroupById(id)
		// console.log(id);
	}
	console.log('done');
}


const removeGroups = async _ => {
	const groupsToExclude = {
		Administrators: true,
		Developers: true,
		Everyone: true
	}
	const groups = await getAllGroups();
	console.log(groups);
	const filtereds = groups.filter(el => !groupsToExclude[el.LoginName])
	console.log(filtereds);
	const ids = filtereds.map(el => el.Id);
	console.log(ids);
	// removeGroupsByIds(ids)
}
// removeGroups()


const uploadFileJSOM = content => {
	const webUrl = '/test/spx';
	const listUrl = 'Files';
	const filename = 'add1.png';
	const clientContext = new SP.ClientContext(webUrl);
	const list = clientContext.get_web().get_lists().getByTitle(listUrl);
	const files = list.get_rootFolder().get_files();
	const fileCreationInfo = new SP.FileCreationInformation;

	fileCreationInfo.set_url(`${webUrl}/${listUrl}/${filename}`);
	fileCreationInfo.set_content(content)

	const file = files.add(fileCreationInfo);
	clientContext.load(file);
	clientContext.executeQueryAsync(console.log, console.error);
}

const uploadFile = async ({ webUrl = '/test/spx', listGUID = '8f0fca61-640a-422d-885e-71e7109dce18', filename = 'add1.png', blob }) => {
	let founds;
	const inputs = [];
	const requiredInputs = {
		__REQUESTDIGEST: true,
		__VIEWSTATE: true,
		__EVENTTARGET: true,
		__EVENTVALIDATION: true,
	}
	const res = await axios(`${webUrl}/_layouts/15/Upload.aspx?List={${listGUID}}`);
	const formMatches = res.data.match(/<form(\w|\W)*<\/form>/);
	const inputRE = /<input[^<]*\/>/g;
	while (founds = inputRE.exec(formMatches)) {
		let item = founds[0];
		const id = item.match(/id=\"([^\"]+)\"/)[1];
		if (requiredInputs[id]) {
			if (id === '__EVENTTARGET') item = item.replace(/value="[^\"]*"/, 'value="ctl00$PlaceHolderMain$ctl03$RptControls$btnOK"');
			inputs.push(item);
		}
	}
	const form = window.document.createElement('form');
	form.innerHTML = inputs.join('');
	const formData = new FormData(form);
	formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', blob, filename);
	await axios({
		url: `${webUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
		method: 'POST',
		data: formData,
	}).catch(res => {
		if (/ctl00_PlaceHolderMain_LabelMessage/i.test(res.data)) throw new Error(res.data.match(/ctl00_PlaceHolderMain_LabelMessage">([^<]+)<\/span>/)[1])
	})
	console.log('done')
}

const upload = async _ => {
	const blob = await spx('test/spx').file('Files/add.png').get({ asBlob: true });
	uploadFile({ blob });
}

// upload()

const getFileBase64 = async _ => {
	const newFile = await spx('test/spx').list('Files').file('add.png').get({ asBlob: true });
	console.log(newFile);
	const content = await new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = e => resolve(e.srcElement.result);
		reader.readAsText(newFile);
	})
	console.log(content);
}

// getFileBase64()


const createItem = async _ => {
	const clientContext = new SP.ClientContext('/test/spx');
	const list = clientContext.get_web().get_lists().getByTitle('Items');
	const creationInfo = new SP.ListItemCreationInformation();
	creationInfo.set_folderUrl('/test/spx/Lists/Items/a');
	const item = list.addItem(creationInfo);
	item.set_item('Title', 'hi');
	item.update();
	clientContext.load(item);
	clientContext.executeQueryAsync(_ => {
		console.log(item);
	})
}
// createItem()

const getPaged = el => new Promise((resolve, reject) => {
	const clientContext = new SP.ClientContext('/test/spx');
	const camlQuery = new SP.CamlQuery();
	if (el) {
		const position = new SP.ListItemCollectionPosition();
		const pagingInfo = `Paged=TRUE&p_ID=${el.ID}&p_sort1=${el.sort1}&p_sort2=${el.sort2}&p_sort3=${el.sort3}`;
		position.set_pagingInfo(pagingInfo);
		camlQuery.set_listItemCollectionPosition(position);
	}
	camlQuery.set_viewXml(getCamlView({ Limit: 2, OrderBy: ['sort1', 'sort2', 'sort3>'] }))
	const itemsSP = clientContext.get_web().get_lists().getByTitle('PagerSort').getItems(camlQuery);
	const items = clientContext.loadQuery(itemsSP, 'Include(ID,sort1,sort2,sort3)');
	clientContext.executeQueryAsync(_ => resolve(items.map(el => el.get_fieldValues())), console.log)
})

const getPages = async _ => {
	const p1 = await getPaged();
	console.log(p1);
	const last = p1[p1.length - 1];
	const p2 = await getPaged(last)
	console.log(p2);
}

// getPages()

const getPagesSPX = async _ => {
	const list = spx('test/spx').list('PagerSort');
	const p1 = await list.item({
		Limit: 2,
		OrderBy: 'ID>',
		// OrderBy: ['sort1', 'sort2', 'sort3>'],
		Page: {}
	}).get({ view: ['ID', 'sort1', 'sort2', 'sort3'] });
	console.log(p1);
	const last = p1[p1.length - 1];
	const p2 = await list.item({
		Limit: 2,
		OrderBy: 'ID>',
		// OrderBy: ['sort1', 'sort2', 'sort3>'],
		Page: {
			Id: last.ID,
			// Columns: {
			// 	sort1: last.sort1,
			// 	sort2: last.sort2,
			// 	sort3: last.sort3
			// }
		}
	}).get({ view: ['ID', 'sort1', 'sort2', 'sort3'] });
	console.log(p2);
}

// getPagesSPX()

const destructor = x => console.log(...x)
// console.log(destructor());

const copyPostImages = async _ => {
	const feedsLibrary = spx('Lenta').library('NewsMedia');
	const postsLibrary = spx('test/spx').library('Files');
	const postsPromise = postsLibrary.item({ Scope: 'items' }).get({ view: ['FileLeafRef'], groupBy: ['FileLeafRef'] });

	const feedPrmoise = feedsLibrary.item().get({ view: ['FileLeafRef', 'FileRef'] })
	const [feeds, posts] = await Promise.all([feedPrmoise, postsPromise])

	console.log(posts);
	console.log(feeds);
	const feedsFiltered = feeds.filter(el => !posts[el.FileLeafRef]);
	console.log(feedsFiltered);
	await proceedImages(postsLibrary)([feedsFiltered[0]]);
	console.log('done');
	// proceedImages(postsLibrary)(newsFiltered.concat(feedsFiltered));
}

const proceedImages = library => async items => {
	let counter = items.length;
	const body = document.getElementsByTagName('body')[0];
	for (const item of items) {
		body.innerHTML = `<div>${counter--}<div>`;
		const url = `http://aura.dme.aero.corp${item.FileRef}`;
		const source = await spx('Lenta').library('NewsMedia').file(item.FileLeafRef).get({ asBlob: true });
		const reader = new FileReader();
		reader.readAsDataURL(source);
		const base64 = await new Promise((resolve, reject) => {
			reader.onloadend = _ => resolve(reader.result)
		})
		console.log(source);
		// console.log(base64);
		await library.file({
			Url: `/test/spx/Files${item.FileLeafRef}`,
			Content: base64.split('base64,')[1].split('base64,')[1],
			// Columns: { Title: 'hi' }
		}).create()
	}
}


const convertImage = ({ url, width }) => new Promise((resolve, reject) => {
	const image = new Image();
	image.src = url;
	image.onload = _ => {
		const size = resizer({ w: image.width, h: image.height, maxWidth: width || 1920 });
		const canvas = document.createElement('canvas');
		canvas.width = size.width;
		canvas.height = size.height;
		resolve(canvas)
		// pica().resize(image, canvas).then(resolve)
	}
	image.onerror = err => {
		console.log(err, url);
		resolve()
	}
})

const resizer = ({ w, h, maxWidth }) => {
	const k = maxWidth / w;
	const size = {
		width: w,
		height: h
	}
	if (k < 1) {
		size.width = maxWidth;
		size.height *= k;
	}
	return size
}
// copyPostImages()

const uploadTestImage = async  _ => {
	const name = '1c905841d5dd82b76894d42cbd3ab9140A09620803741846129.jpeg';
	const url = `/Lenta/NewsMedia/${name}`;
	const blob = await spx('Lenta').library('NewsMedia').file(url).get({ asBlob: true });

	await spx('test/spx').library('Files').file({
		// Url:name,
		Url: 'binary.txt',
		Content: 'hi',
		// Content: await blobToArrayBuffer(blob),
		Columns: { Title: 'hi1' }
	}).create({ silent: true })
}

const a = false;
if (a) throw new Error()
// uploadTestImage();
// spx('Lenta').library('NewsMedia').file('/Lenta/NewsMedia/1c905841d5dd82b76894d42cbd3ab9140A09620803741846129.jpeg').get().then(log)
// spx('test/spx').library('Files').file('/test/spx/Files/1c905841d5dd82b76894d42cbd3ab9140A09620803741846129.jpeg').get().then(log)