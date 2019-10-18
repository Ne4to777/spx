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
	'AllowMembersEditMembership',
	'AllowRequestToJoinLeave',
	'AutoAcceptRequestToJoinLeave',
	'Description',
	'Id',
	'IsHiddenInUI',
	'LoginName',
	'OnlyAllowMembersViewMembership',
	'OwnerTitle',
	'PrincipalType',
	'RequestToJoinLeaveEmailSetting',
	'Title'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = web()
const groupName = 'testGroup'
const groupNameAnother = 'testGroupAnother'
const groupDescription = 'test'
const groupDescriptionAnother = 'testAnother'

const crud = async () => {
	await rootWeb
		.group(groupName)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)

	const newGroup = await assertObjectProps('new group')(rootWeb.group(groupName).create())
	assert(`Title is not a "${groupName}"`)(newGroup.Title === groupName)
	const updatedGroup = await assertObjectProps('new group')(rootWeb.group({
		Title: groupName,
		Description: groupDescription
	}).update())
	assert(`Description is not a "${groupDescription}"`)(updatedGroup.Description === groupDescription)
	await rootWeb.group(groupName).delete({ noRecycle: true })
}

const crudCollection = async () => {
	await rootWeb
		.group([groupName, groupNameAnother])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newGroups = await assertCollectionProps('new group')(
		rootWeb.group([groupName, groupNameAnother]).create()
	)
	assert(`Title is not a "${groupName}"`)(newGroups[0].Title === groupName)
	assert(`Title is not a "${groupNameAnother}"`)(newGroups[1].Title === groupNameAnother)
	const updateGroups = await assertCollectionProps('new group')(
		rootWeb.group([{
			Title: groupName,
			Description: groupDescription
		}, {
			Title: groupNameAnother,
			Description: groupDescriptionAnother
		}]).update()
	)
	assert(`Description is not a "${groupDescription}"`)(updateGroups[0].Description === groupDescription)
	assert(`Description is not a "${groupDescriptionAnother}"`)(updateGroups[1].Description === groupDescriptionAnother)
	await rootWeb.group([groupName, groupNameAnother]).delete({ noRecycle: true })
}

export default {
	get: () => testWrapper('group GET')([
		() => assertObjectProps('group')(rootWeb.group('Test').get()),
		() => assertCollectionProps('group collection')(rootWeb.group().get()),
		() => assertCollectionProps('group collection array')(rootWeb.group(['Test']).get()),
	]),
	crud: () => testWrapper('group CRUD')([crud]),
	crudCollection: () => testWrapper('group CRUD Collection')([crudCollection]),
	all() {
		testWrapper('group ALL')([
			this.get,
			this.crud,
			this.crudCollection,
		])
	}
}
