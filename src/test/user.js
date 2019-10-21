/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testWrapper,
	assert,
	identity
} from '../lib/utility'

const PROPS_SP = [
	'Email',
	'Id',
	'IsHiddenInUI',
	'IsSiteAdmin',
	'LoginName',
	'PrincipalType',
	'Title',
	'UserId'
]

const PROPS_NATIVE = [
	'AdjustHijriDays',
	'AltCalendarType',
	'AppAuthor',
	'AppEditor',
	'Attachments',
	'Author',
	'CalendarType',
	'CalendarViewOptions',
	'ContentLanguages',
	'ContentTypeId',
	'Created',
	'Created_x0020_Date',
	'Deleted',
	'Department',
	'EMail',
	'Editor',
	'FSObjType',
	'FileDirRef',
	'FileLeafRef',
	'FileRef',
	'File_x0020_Type',
	'FirstName',
	'FolderChildCount',
	'GUID',
	'ID',
	'InstanceID',
	'IsActive',
	'IsSiteAdmin',
	'ItemChildCount',
	'JobTitle',
	'LastName',
	'Last_x0020_Modified',
	'Locale',
	'MUILanguages',
	'MetaInfo',
	'MobilePhone',
	'Modified',
	'Name',
	'Notes',
	'Office',
	'Order',
	'Picture',
	'ProgId',
	'SPSPictureExchangeSyncState',
	'SPSPicturePlaceholderState',
	'SPSPictureTimestamp',
	'SPSResponsibility',
	'ScopeId',
	'SipAddress',
	'SortBehavior',
	'SyncClientId',
	'Time24',
	'TimeZone',
	'Title',
	'UniqueId',
	'UserEmail',
	'UserInfoHidden',
	'UserName',
	'WebSite',
	'WorkDayEndHour',
	'WorkDayStartHour',
	'WorkDays',
	'WorkPhone',
	'WorkflowInstanceID',
	'WorkflowVersion',
	'owshiddenversion',
	'_CopySource',
	'_HasCopyDestinations',
	'_IsCurrentVersion',
	'_Level',
	'_ModerationComments',
	'_ModerationStatus',
	'_UIVersion',
	'_UIVersionString'
]

const PROPS = [
	'AdmID',
	'AppAuthor',
	'AppEditor',
	'Attachments',
	'Author',
	'Avatar',
	'AvatarBackup',
	'AvatarPos',
	'AvatarPosBackup',
	'BirthDate',
	'ContentTypeId',
	'Created',
	'Created_x0020_Date',
	'DurationStay',
	'Editor',
	'Education',
	'Email',
	'ExtraEmails',
	'ExtraPhones',
	'FSObjType',
	'FileDirRef',
	'FileLeafRef',
	'FileRef',
	'File_x0020_Type',
	'FolderChildCount',
	'FullPath',
	'GUID',
	'Gender',
	'Hobbies',
	'ID',
	'InstanceID',
	'ItemChildCount',
	'Jobs',
	'KeyField',
	'Knowledges',
	'Languages',
	'LastAccessTS',
	'LastPage',
	'Last_x0020_Modified',
	'Login',
	'MetaID',
	'MetaInfo',
	'Modified',
	'OfficeAddr',
	'OpenPage',
	'Order',
	'PC',
	'Path',
	'PersonID',
	'PhoneInt',
	'PhoneMobile',
	'Position',
	'ProgId',
	'Room',
	'ScopeId',
	'ShortPath',
	'SortBehavior',
	'Space',
	'SyncClientId',
	'Title',
	'UniqueId',
	'WorkflowInstanceID',
	'WorkflowVersion',
	'checksum',
	'deleted',
	'isActive',
	'owshiddenversion',
	'uid',
	'_CopySource',
	'_HasCopyDestinations',
	'_IsCurrentVersion',
	'_Level',
	'_ModerationComments',
	'_ModerationStatus',
	'_UIVersion',
	'_UIVersionString',
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const assertObjectPropsSP = assertObject(PROPS_SP)
const assertCollectionPropsSP = assertCollection(PROPS_SP)

const assertObjectPropsNative = assertObject(PROPS_NATIVE)
const assertCollectionPropsNative = assertCollection(PROPS_NATIVE)

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

	const newGroup = await assertObjectPropsSP('new group')(rootWeb.group(groupName).user(userName).create())
	assert(`LoginName is not a "${userName}"`)(newGroup.LoginName === userName)
	await rootWeb.group(groupName).user(userName).delete({ noRecycle: true })
}

const crudCollection = async () => {
	await rootWeb
		.group(groupName)
		.user([userName, userNameAnother])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newGroups = await assertCollectionPropsSP('new group')(
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
		() => assertObjectPropsSP('user current sp')(rootWeb.userSP().get()),
		() => assertObjectProps('user current')(rootWeb.user().get()),
		() => assertCollectionPropsNative('user title sp')(rootWeb.userSP('Алексеев Алексей Сергеевич').get()),
		() => assertCollectionProps('user title')(rootWeb.user('Алексеев Алексей Сергеевич').get()),
		() => assertObjectPropsNative('user login sp')(rootWeb.userSP('asalekseev').get()),
		() => assertObjectProps('user login')(rootWeb.user('asalekseev').get()),
		() => assertObjectPropsNative('user mail sp')(rootWeb.userSP('AlekseyAlekseew@dme.ru').get()),
		() => assertObjectProps('user mail')(rootWeb.user('AlekseyAlekseew@dme.ru').get()),
		() => assertObjectPropsNative('user id sp')(rootWeb.userSP(10842).get()),
		() => assertObjectProps('user id')(rootWeb.user(10842).get()),
	]),
	// crud: () => testWrapper('user CRUD')([crud]),
	// crudCollection: () => testWrapper('user CRUD Collection')([crudCollection]),
	// all() {
	// 	testWrapper('user ALL')([
	// 		this.get,
	// 		this.crud,
	// 		this.crudCollection,
	// 	])
	// }
}
