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

console.log(1)

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

window.run = async () => {
	const blob1 = new Blob(['hi'], { type: 'text/plain' })
	const blob2 = new Blob(['hi1'], { type: 'text/plain' })
	const folders = ['a/b/c', 'a', 'b/c/d', 'b', 'a/b/d']
	const web = spx('test/spx')
	const list = web.list('Folders')
	const library = web.library('Documents')
	// list
	// 	.item([{ Title: 'hi', Folder: 'a/b/c' }])
	// 	.create()
	// 	.then(console.log)
	// return
	// const item = { Columns: { Title: 'new item' } }
	const file = { Columns: { Title: 'new item' } }
	await library
		// .item([{ ...item, Folder: 'a' }, { ...item, Folder: 'b' }])
		.file([{
			// ...file,
			Folder: 'a',
			Url: 'test.txt',
			Content: blob1,
		}
			// , {
			// 	...file,
			// 	Folder: 'b',
			// 	Url: 'test1.txt',
			// 	Content: blob2
			// }
		])
		.create()
	console.log('done')
}


const getTermStore = clientContext => SP
	.Taxonomy
	.TaxonomySession
	.getTaxonomySession(clientContext)
	.getDefaultKeywordsTermStore()

const getTermSet = clientContext => getTermStore(clientContext).get_keywordsTermSet()


const getAllTerms = clientContext => getTermSet(clientContext).getAllTerms()

const getSPObject = (clientContext, elementUrl) => getAllTerms(clientContext).getByName(elementUrl)

const clientContext = new SP.ClientContext('http://localhost:3000')

const taxonomySession = SP.Taxonomy
	.TaxonomySession
	.getTaxonomySession(clientContext)

// clientContext.load(taxonomySession)

const termStore = taxonomySession.getDefaultSiteCollectionTermStore()
// clientContext.load(termStore)

const hashTags = termStore.get_hashTagsTermSet()
clientContext.load(hashTags)

// const keywords = termStore.get_keywordsTermSet()
// clientContext.load(keywords)

// const groups = termStore.get_groups()
// clientContext.load(groups)

// const group = groups.getByName('Site Collection - sharepoint.local')
// clientContext.load(group)

// const termSets = group.get_termSets()
// clientContext.load(termSets)

// const termSet = termSets.getByName('AnotherTermSet')
// clientContext.load(termSet)

// const terms = termSet.get_terms()
// clientContext.load(terms)

clientContext.executeQueryAsync(() => {
	console.log(taxonomySession)
	console.log(termStore)
	console.log(hashTags)
	// console.log(keywords)
	// console.log(groups)
	// console.log(group)
	// console.log(termSets)
	// console.log(termSet)
	// console.log(terms)
}, console.log)


// spx('app-library')
// 	.list('Library')
// 	.item()
// 	.get({ view: ['url'] })
// 	.then(xs => xs
// 		.reduce((acc, x) => JSON
// 			.parse(x.url)
// 			.reduce((ys, url) => {
// 				url
// 					.split('/')
// 					.slice(1, -1)
// 					.reduce((zx, urlSplit) => {
// 						if (!zx[urlSplit]) zx[urlSplit] = {}
// 						return zx[urlSplit]
// 					}, ys)
// 				return ys
// 			}, acc), {}))
// 	.then(console.log)


// console.log([
// 	'/AM/UsersAD',
// 	'/app-list/social/surveys/Surveys',
// 	'/Intellect/SurveyItems',
// 	'/Intellect/Отчеты о командировках',
// 	'/Intellect/TripStories',
// 	'/app-library/Library',
// 	'/Intellect/DME',
// 	'/Intellect/DocLib1',
// 	'/Intellect/DocLib4',
// 	'/app-library/LibraryStorage',
// 	'/lib/StorageLib',
// 	'/app/Blogs',
// 	'/app-list/articles/feed/Posts',
// 	'/app-list/articles/feed/Stories',
// 	'/wikilibrary/wiki/Pages',
// 	'/app-library/ArchiveImages',
// 	'/lib/ArchGov',
// 	'/wikilibrary/wiki/Pages',
// 	'/app-list/EmployeesVeterans',
// 	'/board/BoardList',
// 	'/app-list/social/labs/Projects',
// 	'/app-list/social/labs/Meetings',
// 	'/app-list/social/labs/Lectures',
// 	'/app-list/social/labs/Labs',
// 	'/app-list/social/labs/Materials',
// 	'/app-library/EditorFiles'
// ].reduce((acc, el) => {
// 	el
// 		.split('/')
// 		.slice(1)
// 		.reduce((zx, urlSplit) => {
// 			if (!zx[urlSplit]) zx[urlSplit] = {}
// 			return zx[urlSplit]
// 		}, acc)
// 	return acc
// }, {}))


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

clientContext.load(taxonomySession)
// clientContext.executeQueryAsync(() => {
// 	console.log(taxonomySession)
// 	console.log(termStore)
// 	console.log(hashTags)
// 	console.log(keywords)
// 	console.log(groups)
// 	console.log(group)
// 	console.log(termSets)
// 	console.log(termSet)
// 	console.log(terms)
// })

// const enrollment = [
// 	{ enrolled: 2, grade: 100 },
// 	{ enrolled: 2, grade: 80 },
// 	{ enrolled: 1, grade: 89 }
// ]

const enrollment = lazy.generate(el => ({ enrolled: Math.round(Math.random()) + 1, grade: 100 }))
	.take(100000)
	.toArray()

console.time('imprative')
let totalGrades = 0
let totalStudentsFound = 0
for (let i = 0; i < enrollment.length; i += 1) {
	const student = enrollment[i]
	if (student && student.enrolled > 1) {
		totalGrades += student.grade
		totalStudentsFound += 1
	}
}
console.timeEnd('imprative')
const average1 = totalGrades / totalStudentsFound
// console.log(average1)


console.time('functional1')
const average2 = lazy(enrollment)
	.filter(el => el && el.enrolled > 1)
	.pluck('grade')
	.reduce((acc, el) => ([acc[0] + el, acc[1] += 1]), [0, 0])
	.reduce((acc, el) => acc ? acc / el : el, 0)
console.timeEnd('functional1')
console.log(average2)


console.time('functional2')
const totals2 = enrollment
	.filter(el => el && el.enrolled > 1)
	.map(el => el.grade)
	.reduce((acc, el) => ([acc[0] + el, acc[1] += 1]), [0, 0])
	.reduce((acc, el) => acc ? acc / el : el, 0)
console.timeEnd('functional2')
// const average3 = totals2[0] / totals2[1]
// console.log(totals2)

const U = f => f(f)
const Y = f => U(g => f((...xs) => g(g)(...xs)))


const partial = f => Y(fR => (...xs) => f.length > xs.length ? (...ys) => fR(...xs, ...ys) : f(...xs))


const partialPH = (f, ph = '_') => Y(fPartial => (...xs) => f.length > xs.filter(x => x !== ph).length
	? (...ys) => fPartial(...Y(fMerge => acc => ([x, ...xsTail]) => ([y, ...ysTail]) => x === undefined
		? Array.prototype.concat.apply(acc, y === undefined ? ysTail : [y, ...ysTail])
		: x === ph
			? fMerge(acc.concat(y))(xsTail)(ysTail)
			: fMerge(acc.concat(x))(xsTail)(y === ph ? ysTail : [y, ...ysTail]))([])(xs)(ys))
	: f(...xs))


const sum3 = (a, b, c) => a + b + c

const g = partialPH(sum3)

const assert = input => output => input === output ? undefined : console.error(input, output)

assert(g('a', 'b', 'c'))('abc')
assert(g('a')('b')('c'))('abc')
assert(g('a', 'b')('c'))('abc')
assert(g('a')('b', 'c'))('abc')
assert(g('_', '_', 'a')('_', 'b')('c'))('cba')
assert(g('_', 'a')('_', '_', 'b')('c'))('cab')
assert(g('a')('_', 'b', '_')('_', '_', 'c'))('abc')
assert(g('_', '_', 'a')('_', 'b', '_')('c', '_', '_'))('cba')
assert(g('_', '_', 'c')('_', 'b')('a'))('abc')
assert(g('_', 'b')('_', '_', 'c')('a'))('abc')

console.log(g()()(undefined, undefined)()('_', '_')('a')('b', 'c'))

// const a = ['_', 2, 3, '_', '_']
// const b = [1, 4, '_', 6]
const a = ['_', '_', 'a']
const b = ['_', 'b']
// console.log(mergeArrays([])(a)(b))
