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
	'Email',
	'Id',
	'IsHiddenInUI',
	'IsSiteAdmin',
	'LoginName',
	'PrincipalType',
	'Title',
	'UserId'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = web()
const groupName = 'Test'
const userName = 'DME\\spps_developer'
const userNameAnother = 'DME\\spps_designers'

const crud = async () => {
	await rootWeb
		.group(groupName)
		.user(userName)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)

	const newGroup = await assertObjectProps('new group')(rootWeb.group(groupName).user(userName).create())
	assert(`LoginName is not a "${userName}"`)(newGroup.LoginName === userName)
	await rootWeb.group(groupName).user(userName).delete({ noRecycle: true })
}

const crudCollection = async () => {
	await rootWeb
		.group(groupName)
		.user([userName, userNameAnother])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newGroups = await assertCollectionProps('new group')(
		rootWeb
			.group(groupName)
			.user([userName, userNameAnother])
			.create()
	)
	assert(`LoginName is not a "${userName}"`)(newGroups[0].LoginName === userName)
	assert(`LoginName is not a "${userNameAnother}"`)(newGroups[1].LoginName === userNameAnother)
	await rootWeb
		.group(groupName)
		.user([userName, userNameAnother])
		.delete({ noRecycle: true })
}

export default {
	get: () => testWrapper('user GET')([
		() => assertObjectProps('user')(rootWeb
			.group('Test')
			.user('c:0+.w|s-1-5-21-767579643-3516020610-614104799-43390')
			.get()),
		() => assertCollectionProps('user collection')(rootWeb.group('Test').user().get()),
	]),
	crud: () => testWrapper('user CRUD')([crud]),
	crudCollection: () => testWrapper('user CRUD Collection')([crudCollection]),
	all() {
		testWrapper('user ALL')([
			this.get,
			this.crud,
			this.crudCollection,
		])
	}
}
