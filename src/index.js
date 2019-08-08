/* eslint max-len:0 */
/* eslint no-unused-vars:0 */
/* eslint no-restricted-syntax:0 */
/* eslint import/no-extraneous-dependencies:0 */
import axios from 'axios'
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
	customWebTitle: 'AM',
	customListTitle: 'UsersAD',
	customIdColumn: 'uid',
	customLoginColumn: 'Login',
	customNameColumn: 'Title',
	customEmailColumn: 'Email',
	customQuery: 'Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))'
})

// test()
window.run = async () => {
	const folders = ['a/b/c', 'a', 'b/c/d', 'b', 'a/b/d']
	const list = spx('test/spx').list('Folders')
	spx('test/spx').list('Folders').folder(folders).create()
	return
	const item = { Columns: { Title: 'new item' } }
	await list
		.item([{ ...item, Folder: 'a' }, { ...item, Folder: 'b' }])
		.create({ view: ['ID', 'Title', 'FileDirRef'] })
	console.log('done')
}


