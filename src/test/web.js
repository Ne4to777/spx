/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testIsOk,
	assert,
	map,
	identity
} from '../lib/utility'

const SITE_PROPS = [
	'AllowCreateDeclarativeWorkflow',
	'AllowDesigner',
	'AllowMasterPageEditing',
	'AllowRevertFromTemplate',
	'AllowSaveDeclarativeWorkflowAsTemplate',
	'AllowSavePublishDeclarativeWorkflow',
	'AllowSelfServiceUpgrade',
	'AllowSelfServiceUpgradeEvaluation',
	'AuditLogTrimmingRetention',
	'CompatibilityLevel',
	'Id',
	'LockIssue',
	'MaxItemsPerThrottledOperation',
	'PrimaryUri',
	'ReadOnly',
	'RequiredDesignerVersion',
	'ServerRelativeUrl',
	'ShareByLinkEnabled',
	'ShowUrlStructure',
	'TrimAuditLog',
	'UIVersionConfigurationEnabled',
	'UpgradeReminderDate',
	'Upgrading',
	'Url'
]

const WEB_TEMPLATE_PROPS = [
	'Description',
	'DisplayCategory',
	'Id',
	'ImageUrl',
	'IsHidden',
	'IsRootWebOnly',
	'IsSubWebOnly',
	'Lcid',
	'Name',
	'Title'
]

const LIST_TEMPLATE_PROPS = [
	'AllowsFolderCreation',
	'BaseType',
	'Description',
	'FeatureId',
	'ImageUrl',
	'InternalName',
	'IsCustomTemplate',
	'ListTemplateTypeKind',
	'Name',
	'OnQuickLaunch',
	'Unique'
]


const PROPS = [
	'AllowRssFeeds',
	'AlternateCssUrl',
	'AppInstanceId',
	'Configuration',
	'Created',
	'CustomMasterUrl',
	'Description',
	'DocumentLibraryCalloutOfficeWebAppPreviewersDisabled',
	'EnableMinimalDownload',
	'Id',
	'Language',
	'LastItemModifiedDate',
	'MasterUrl',
	'QuickLaunchEnabled',
	'RecycleBinEnabled',
	'ServerRelativeUrl',
	'SiteLogoUrl',
	'SyndicationEnabled',
	'Title',
	'TreeViewEnabled',
	'UIVersion',
	'UIVersionConfigurationEnabled',
	'Url',
	'WebTemplate'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const crud = async () => {
	const name = 'test/spx/createdWebSingle1'
	await web(name)
		.delete({ noRecycle: true, silentErrors: true })
		.catch(identity)
	const newWeb = await assertObjectProps('new web')(web({
		Url: name,
		Description: 'Default Aura Web Template'
	}).create())
	assert('Description is not a "Default Aura Web Template"')(newWeb.Description === 'Default Aura Web Template')
	const updatedWeb = await assertObjectProps('updated web')(
		web({ Url: name, Description: 'Test description' }).update()
	)
	assert('Description is not a "Test description"')(updatedWeb.Description === 'Test description')
	await web(name).delete({ noRecycle: true })
}
const crudCollection = async () => {
	const name1 = 'test/spx/createdWebMulti3'
	const name2 = 'test/spx/createdWebMultiAnother3'
	await web([name1, name2])
		.delete({ noRecycle: true })
		.catch(identity)
	const newWebs = await assertCollectionProps('new webs')(
		web([{
			Url: name1,
			Description: 'Default Aura Web Template'
		}, {
			Url: name2,
			Description: 'Default Aura Web Template'
		}]).create()
	)
	map(newWeb => assert('Description is not a "Default Aura Web Template"')(
		newWeb.Description === 'Default Aura Web Template'
	))(newWebs)
	const updatedWebs = await assertCollectionProps('updated web')(
		web([
			{ Url: name1, Description: 'Test description' },
			{ Url: name2, Description: 'Test description another' }
		]).update()
	)
	assert('Description is not a "Test description"')(updatedWebs[0].Description === 'Test description')
	assert('Description is not a "Test description another"')(updatedWebs[1].Description === 'Test description another')
	await web([name1, name2]).delete({ noRecycle: true })
}

const doesUserHavePermissions = async () => {
	const has = await web('test/spx/testWeb').doesUserHavePermissions()
	assert('user has wrong permissions for web')(has)
}

export default () => Promise.all([
	// assertObject(SITE_PROPS)('site')(web().getSite()),
	// assertCollection(WEB_TEMPLATE_PROPS)('web template')(web().getWebTemplates()),
	// assertCollection(LIST_TEMPLATE_PROPS)('custom lists template')(web().getCustomListTemplates()),
	// assertObjectProps('root web')(web().get()),
	// assertCollectionProps('root webs')(web('/').get()),
	// assertObjectProps('web')(web('test/spx/testWeb').get()),
	// assertCollectionProps('test/spx/testWeb')(web('test/spx/').get()),
	// assertCollectionProps('web')(web(['test/spx/testWeb']).get()),
	// assertObjectProps('web')(web({ Url: 'test/spx/testWeb' }).get()),
	// assertCollectionProps('web')(web([{ Url: 'test/spx/testWeb' }]).get()),
	// assertCollectionProps('web')(web(['test/spx/testWeb', 'test/spx/testWebAnother']).get()),
	// assertCollectionProps('web')(web([{ Url: 'test/spx/testWeb' }, { Url: 'test/spx/testWebAnother' }]).get()),
	// doesUserHavePermissions(),
	// crud(),
	crudCollection()
]).then(testIsOk('web'))
