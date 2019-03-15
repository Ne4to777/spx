import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, map } from './../utility'

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
];

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const crud = async _ => site('test/spx/createdWebSingle').delete({ noRecycle: true, silentErrors: true }).catch(async err => {
  const newWeb = await assertObjectProps('new web')(site('test/spx/createdWebSingle').create());
  assert(`Description is not a "Default Aura Web Template"`)(newWeb.Description === 'Default Aura Web Template')
  const updatedWeb = await assertObjectProps('updated web')(site({ Url: 'test/spx/createdWebSingle', Description: 'Test description' }).update());
  assert(`Description is not a "Test description"`)(updatedWeb.Description === 'Test description')
  await site('test/spx/createdWebSingle').delete({ noRecycle: true });
});
const crudCollection = async _ => site(['test/spx/createdWebMulti', 'test/spx/createdWebMultiAnother']).delete({ noRecycle: true, silentErrors: true }).catch(async err => {
  const newWebs = await assertCollectionProps('new webs')(site(['test/spx/createdWebMulti', 'test/spx/createdWebMultiAnother']).create());
  map(newWeb => assert(`Description is not a "Default Aura Web Template"`)(newWeb.Description === 'Default Aura Web Template'))(newWebs)
  const updatedWebs = await assertCollectionProps('updated web')(site([
    { Url: 'test/spx/createdWebMulti', Description: 'Test description' },
    { Url: 'test/spx/createdWebMultiAnother', Description: 'Test description another' }
  ]).update());
  assert(`Description is not a "Test description"`)(updatedWebs[0].Description === 'Test description')
  assert(`Description is not a "Test description another"`)(updatedWebs[1].Description === 'Test description another')
  await site(['test/spx/createdWebMulti', 'test/spx/createdWebMultiAnother']).delete({ noRecycle: true });
});

const doesUserHavePermissions = async  _ => {
  const has = await site('test/spx/testWeb').doesUserHavePermissions();
  assert(`user has wrong permissions for web`)(has)
}

// console.log(crud());
export default _ => Promise.all([
  assertObjectProps('root web')(site().get()),
  assertCollectionProps('root webs')(site('/').get()),
  assertObjectProps('web')(site('test/spx/testWeb').get()),
  assertCollectionProps('test/spx/testWeb')(site('test/spx/').get()),
  assertCollectionProps('web')(site(['test/spx/testWeb']).get()),
  assertObjectProps('web')(site({ Url: 'test/spx/testWeb' }).get()),
  assertCollectionProps('web')(site([{ Url: 'test/spx/testWeb' }]).get()),
  assertCollectionProps('web')(site(['test/spx/testWeb', 'test/spx/testWebAnother']).get()),
  assertCollectionProps('web')(site([{ Url: 'test/spx/testWeb' }, { Url: 'test/spx/testWebAnother' }]).get()),
  doesUserHavePermissions(),
  // crud(),
  // crudCollection()
]).then(testIsOk('web'))
