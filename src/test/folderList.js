import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, filter, isObjectFilled } from './../utility';

const PROPS = [
  'ItemCount',
  'Name',
  'ServerRelativeUrl',
  'WelcomePage'
];

const ITEM_PROPS = [
  'AppAuthor',
  'AppEditor',
  'Attachments',
  'Author',
  'ContentTypeId',
  'Created',
  'Created_x0020_Date',
  'Editor',
  'FSObjType',
  'FileDirRef',
  'FileLeafRef',
  'FileRef',
  'File_x0020_Type',
  'FolderChildCount',
  'GUID',
  'ID',
  'InstanceID',
  'ItemChildCount',
  'Last_x0020_Modified',
  'MetaInfo',
  'Modified',
  'Order',
  'ProgId',
  'ScopeId',
  'SortBehavior',
  'SyncClientId',
  'TaxCatchAll',
  'Title',
  'Title1',
  'TranslationStateTermInformation',
  'UniqueId',
  'WorkflowInstanceID',
  'WorkflowVersion',
  'group',
  'lookup',
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

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const assertObjectItemProps = assertObject(ITEM_PROPS);
const assertCollectionItemProps = assertCollection(ITEM_PROPS);

const rootWebList = site().list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2');
const workingWebList = site('test/spx').list('Test');

const crud = async _ => workingWebList.folder('a/singleFolder').delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolder = await assertObjectProps('new folder')(workingWebList.folder('a/singleFolder').create({ silent: true }));
  assert(`Name is not a "singleFolder"`)(newFolder.Name === 'singleFolder');
  const updatedFolder = await assertObjectProps('new folder')(workingWebList.folder({ Url: 'a/singleFolder', Title: 'updated folder' }).update({ silent: true }));
  assert(`Name is not a "singleFolder"`)(updatedFolder.Name === 'singleFolder');
  await workingWebList.folder('a/singleFolder').delete({ noRecycle: true, silent: true });
});

const crudAsItem = async _ => workingWebList.folder('a/singleFolderAsItem').delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolder = await assertObjectItemProps('new folder')(workingWebList.folder('a/singleFolderAsItem').create({ silent: true, asItem: true }));
  assert(`Title is not a "singleFolderAsItem"`)(newFolder.Title === 'singleFolderAsItem');
  const updatedFolder = await assertObjectItemProps('new folder')(workingWebList.folder({ Url: 'a/singleFolderAsItem', Title: 'updated folder' }).update({ silent: true, asItem: true }));
  assert(`Title is not a "updated folder"`)(updatedFolder.Title === 'updated folder');
  await workingWebList.folder('a/singleFolderAsItem').delete({ noRecycle: true, silent: true });
});

const crudCollection = async _ => workingWebList.folder(['a/multiFolder', 'a/multiFolderAnother']).delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolders = await assertCollectionProps('new folder')(workingWebList.folder(['a/multiFolder', 'a/multiFolderAnother']).create({ silent: true }));
  assert(`Name is not a "multiFolder"`)(newFolders[0].Name === 'multiFolder');
  assert(`Name is not a "multiFolderAnother"`)(newFolders[1].Name === 'multiFolderAnother');
  const updatedFolder = await assertCollectionProps('new folder')(workingWebList.folder([
    { Url: 'a/multiFolder', Title: 'updated folder' },
    { Url: 'a/multiFolderAnother', Title: 'updated folder another' }
  ]).update({ silent: true }));
  assert(`Name is not a "multiFolder"`)(updatedFolder[0].Name === 'multiFolder');
  assert(`Name is not a "multiFolderAnother"`)(updatedFolder[1].Name === 'multiFolderAnother');
  await workingWebList.folder(['a/multiFolder', 'a/multiFolderAnother']).delete({ noRecycle: true, silent: true });
});

const crudCollectionAsItem = async _ => workingWebList.folder(['a/multiFolderAsItem', 'a/multiFolderAsItemAnother']).delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolders = await assertCollectionItemProps('new folder')(workingWebList.folder([
    { Url: 'a/multiFolderAsItem', Title: 'new folder' },
    { Url: 'a/multiFolderAsItemAnother', Title: 'new folder another' }]).create({ silent: true, asItem: true }));
  assert(`Title is not a "new folder"`)(newFolders[0].Title === 'new folder');
  assert(`Title is not a "new folder another"`)(newFolders[1].Title === 'new folder another');
  const updatedFolders = await assertCollectionItemProps('new folder')(workingWebList.folder([
    { Url: 'a/multiFolderAsItem', Title: 'updated folder' },
    { Url: 'a/multiFolderAsItemAnother', Title: 'updated folder another' }
  ]).update({ silent: true, asItem: true }));
  assert(`Title is not a "updated folder"`)(updatedFolders[0].Title === 'updated folder');
  assert(`Title is not a "updated folder another"`)(updatedFolders[1].Title === 'updated folder another');
  await workingWebList.folder(['a/multiFolderAsItem', 'a/multiFolderAsItemAnother']).delete({ noRecycle: true, silent: true });
});

export default _ => Promise.all([
  assertObjectProps('root web list folder')(rootWebList.folder().get()),
  assertCollectionProps('root web list folder')(rootWebList.folder('/').get()),
  assertObjectProps('web root list folder')(workingWebList.folder().get()),
  assertObjectProps('web a list folder')(workingWebList.folder('a').get()),
  assertCollectionProps('web a list folder')(workingWebList.folder('a/').get()),
  assertCollectionProps('web a, c list folder')(workingWebList.folder(['a', 'c']).get()),

  assertObjectItemProps('web a list folder')(workingWebList.folder('a').get({ asItem: true })),
  assertCollectionItemProps('web a list folder')(workingWebList.folder('a/').get({ asItem: true }).then(filter(isObjectFilled))),
  assertCollectionItemProps('web a, c list folder')(workingWebList.folder(['a', 'c']).get({ asItem: true })),

  crud(),
  crudCollection(),
  crudAsItem(),
  crudCollectionAsItem()
]).then(testIsOk('folderList'))