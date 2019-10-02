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
	// customWebTitle: 'AM',
	// customListTitle: 'UsersAD',
	// customIdColumn: 'uid',
	// customLoginColumn: 'Login',
	// customNameColumn: 'Title',
	// customEmailColumn: 'Email',
	// customQuery: 'Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))'
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

clientContext.executeQueryAsync(() => {
	console.log(taxonomySession)
	console.log(termStores)
	console.log(termStore)
	console.log(hashTags)
	console.log(keywords)
	console.log(orphanedTags)
	console.log(groups)
	console.log(group)
	console.log(termSets)
	console.log(termSet)
	console.log(terms)
	console.log(term)
}, (sender, args) => console.error(args.get_message()))


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
