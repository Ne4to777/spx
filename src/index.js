/* eslint max-len:0 */
/* eslint no-unused-vars:0 */
/* eslint no-restricted-syntax:0 */
/* eslint import/no-extraneous-dependencies:0 */
import axios from 'axios'
import lazy from 'lazy.js'
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
	typeOf,
	shiftSlash,
	isObject,
	isArray,
	popSlash,
	urlSplit,
	pipe,
	isString,
	fix,
	isObjectFilled
} from './lib/utility'
import * as cache from './lib/cache'

window.test = test

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

// test()

// const getTermStore = clientContext => SP
// 	.Taxonomy
// 	.TaxonomySession
// 	.getTaxonomySession(clientContext)
// 	.getDefaultKeywordsTermStore()

// const getTermSet = clientContext => getTermStore(clientContext).get_keywordsTermSet()


// const getAllTerms = clientContext => getTermSet(clientContext).getAllTerms()

// const getSPObject = (clientContext, elementUrl) => getAllTerms(clientContext).getByName(elementUrl)


const clientContext = new SP.ClientContext('/')

const taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext)

clientContext.load(taxonomySession)

const termStores = taxonomySession.get_termStores()
clientContext.load(termStores)

// const termStore = termStores.getByName('Managed Metadata Service')
const termStore = taxonomySession.getDefaultSiteCollectionTermStore()
clientContext.load(termStore)

const hashTags = termStore.get_hashTagsTermSet()
clientContext.load(hashTags)

// const keywords = termStore.get_keywordsTermSet()
const keywords = taxonomySession.getDefaultKeywordsTermStore().get_keywordsTermSet()
clientContext.load(keywords)

const orphanedTags = termStore.get_orphanedTermsTermSet()
clientContext.load(orphanedTags)

const groups = termStore.get_groups()
clientContext.load(groups)

const group = groups.getByName('Test')
clientContext.load(group)

const termSets = group.get_termSets()
clientContext.load(termSets)

const termSet = termSets.getByName('Test')
clientContext.load(termSet)

const terms = termSet.get_terms()
clientContext.load(terms)

const term = terms.getByName('test')
clientContext.load(term)

// clientContext.executeQueryAsync(() => {
// 	console.log(taxonomySession)
// 	console.log(termStores)
// 	console.log(termStore)
// 	console.log(hashTags)
// 	console.log(keywords)
// 	console.log(orphanedTags)
// 	console.log(groups)
// 	console.log(group)
// 	console.log(termSets)
// 	console.log(termSet)
// 	console.log(terms)
// 	console.log(term)
// }, (sender, args) => console.error(args.get_message()))


window.setListsNowCrawlByWeb = webUrl => spx(webUrl)
	.list()
	.get({
		// view: ['NoCrawl', 'Title']
	})
	.then(log)
	.then(lists => lists
		.filter(el => !el.NoCrawl)
		.map(list => ({
			Title: list.Title,
			NoCrawl: true
		})))
	.then(log)
	.then(listsToUpdate => listsToUpdate.map(listToUpdate => spx(webUrl)
		.list(listToUpdate)
		.update().catch(console.error)))

window.setModStat = () => {
	const c = new SP.ClientContext('/test')
	const list = c.get_web().get_lists().getByTitle('Terms')
	const listItem = list.getItemById(1)
	listItem.set_item('_ModerationStatus', 1)
	listItem.update()
	c.load(listItem)
	c.executeQueryAsync(() => {
		console.log('done')
	}, (s, err) => {
		console.log(err)
	})
}

window.checkPermissions = async () => {
	const inputs = []
	const requiredInputs = {
		__VIEWSTATE: true,
		__EVENTVALIDATION: true,
		__REQUESTDIGEST: true
	}

	const listForms = await axios.get('/test/_layouts/15/chkperm.aspx')
	const listFormMatches = listForms.data.match(/<form[(\w|\W)]*<\/form>/)
	const inputRE = /<input[^<]*\/>/g
	let founds = inputRE.exec(listFormMatches)

	while (founds) {
		const item = founds[0]
		const id = item.match(/id="([^"]+)"/)[1]
		if (requiredInputs[id]) {
			inputs.push(item)
		}
		founds = inputRE.exec(listFormMatches)
	}

	const form = window.document.createElement('form')
	form.innerHTML = inputs.join('')
	const formData = new FormData(form)

	formData.append('__EVENTTARGET', 'ctl00$PlaceHolderMain$buttonSectionMain$RptControls$btnCheckPerm')
	formData.append('ctl00$PlaceHolderMain$ctl00$ctl02$peoplePicker', JSON.stringify([{
		Key: 'i:0#.w|dme\\asalekseev'
	}]))

	const response = await axios({
		url: '/test/_layouts/15/chkperm.aspx?obj=%7B00A96C5E%2D6A1E%2D4511%2D81EE%2DA468BEF85833%7D%2CLIST',
		method: 'POST',
		data: formData
	})

	console.log(response.data)

	return response
}

window.getAllCrawledLists = async () => {
	const host = 'http://aura.dme.aero.corp'
	const hostRE = new RegExp(host)
	const getCrawledLists = async (webUrls, lists) => {
		// console.log(Object.assign({}, webUrls))
		if (webUrls && webUrls.length) {
			const webs = await spx(webUrls).get({ view: ['Url'] }).then(webs => webs.filter(web => hostRE.test(web.Url)).map(web => web.Url.split(host)[1]))
			// console.log(webs)

			for (let i = 0; i < webs.length; i += 1) {
				const webUrl = webs[i]
				// console.log(webUrl)

				lists = lists.concat(await getCrawledListsFromWeb(webUrl))
				const res = await getCrawledLists(await getSubwebs(webUrl), lists).catch(_ => 'error')
				if (res === 'error') break
			}
		}
		return lists
	}
	const getCrawledListsFromWeb = async webUrl => {
		const lists = await spx(webUrl).list('/').get({ view: ['Title', 'NoCrawl', 'ParentWebUrl'] })
		// console.log(lists)

		return lists.filter(list => !list.NoCrawl)
	}

	const getSubwebs = webUrl => spx(`${webUrl}/`).get({ view: ['Url'] }).then(webs => webs.filter(web => hostRE.test(web.Url)).map(web => web.Url.split(host)[1]))
	return getCrawledLists('/', [])
}
