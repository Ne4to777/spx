/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testWrapper,
	assert,
	identity
} from '../lib/utility'

const PROPS = [
	'ElapsedTime',
	'Properties',
	'QueryErrors',
	'QueryId',
	'ResultTables',
	'SpellingSuggestion',
	'TriggeredRules',
	'_ObjectType_'
]

const assertObjectProps = assertObject(PROPS)

export default {
	get: () => testWrapper('search GET')([
		() => assertObjectProps('search "test"')(web().search('test').get()),
		() => assertObjectProps('search next')(web().search('test').next()),
		() => assertObjectProps('search previous')(web().search('test').previous()),
		() => assertObjectProps('search previous')(web().search('test').move({ page: 2 })),
	]),
	all() {
		testWrapper('search ALL')([
			this.get,
		])
	}
}
