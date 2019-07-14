import site from '../modules/site'
import {
	assertObject, assertCollection, testIsOk, assert, identity
} from '../lib/utility'

const PROPS = [
	'AllowContentTypes',
	'BaseTemplate',
	'BaseType',
	'ContentTypesEnabled',
	'Created',
	'DefaultContentApprovalWorkflowId',
	'Description',
	'Direction',
	'DocumentTemplateUrl',
	'DraftVersionVisibility',
	'EnableAttachments',
	'EnableFolderCreation',
	'EnableMinorVersions',
	'EnableModeration',
	'EnableVersioning',
	'EntityTypeName',
	'ForceCheckout',
	'HasExternalDataSource',
	'Hidden',
	'Id',
	'ImageUrl',
	'IrmEnabled',
	'IrmExpire',
	'IrmReject',
	'IsApplicationList',
	'IsCatalog',
	'IsPrivate',
	'ItemCount',
	'LastItemDeletedDate',
	'LastItemModifiedDate',
	'ListItemEntityTypeFullName',
	'MajorVersionLimit',
	'MajorWithMinorVersionsLimit',
	'MultipleDataList',
	'NoCrawl',
	'ParentWebUrl',
	'ServerTemplateCanCreateFolders',
	'TemplateFeatureId',
	'Title'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = site()
const workingWeb = site('test/spx')

const crud = async () => {
	await workingWeb
		.list('Single')
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newList = await assertObjectProps('new list')(
		workingWeb.list({ Url: 'Single', Description: 'new list' }).create()
	)
	assert('Description is not a "new list"')(newList.Description === 'new list')
	const updatedList = await assertObjectProps('updated list')(
		workingWeb.list({ Url: 'Single', Description: 'updated list' }).update()
	)
	assert('Description is not a "updated list"')(updatedList.Description === 'updated list')
	await workingWeb.list('Single').delete({ noRecycle: true })
}

const crudCollection = async () => {
	await workingWeb
		.list(['Multi', 'MultiAnother'])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newLists = await assertCollectionProps('new list')(
		workingWeb
			.list([
				{
					Url: 'Multi',
					Description: 'new Multi list'
				},
				{
					Url: 'MultiAnother',
					Description: 'new MultiAnother list'
				}
			])
			.create()
	)
	assert('Description is not a "new Multi list"')(newLists[0].Description === 'new Multi list')
	assert('Description is not a "new MultiAnother list"')(newLists[1].Description === 'new MultiAnother list')
	const updatedLists = await assertCollectionProps('updated list')(
		workingWeb
			.list([
				{
					Url: 'Multi',
					Description: 'updated Multi list'
				},
				{
					Url: 'MultiAnother',
					Description: 'updated MultiAnother list'
				}
			])
			.update()
	)
	assert('Description is not a "updated Multi list"')(updatedLists[0].Description === 'updated Multi list')
	assert('Description is not a "updated MultiAnother list"')(
		updatedLists[1].Description === 'updated MultiAnother list'
	)
	await workingWeb
		.list([
			{
				Url: 'Multi'
			},
			{
				Url: 'MultiAnother'
			}
		])
		.delete({ noRecycle: true })
}

const doesUserHavePermissions = async () => {
	const has = await site('test/spx')
		.list('Test')
		.doesUserHavePermissions()
	assert('user has wrong permissions for list')(has)
}

export default () => Promise.all([
	assertObjectProps('root web list')(rootWeb.list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2').get()),
	assertCollectionProps('root web list')(rootWeb.list('/').get()),
	assertObjectProps('web list')(workingWeb.list('Test').get()),
	assertCollectionProps('web root list')(workingWeb.list('/').get()),
	assertCollectionProps('web Test, TestAnother list')(workingWeb.list(['Test', 'TestAnother']).get()),
	crud(),
	crudCollection(),
	doesUserHavePermissions()
]).then(testIsOk('list'))
