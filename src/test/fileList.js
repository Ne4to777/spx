import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, filter, isObjectFilled } from './../lib/utility';

const PROPS = [
  'CheckInComment',
  'CheckOutType',
  'ContentTag',
  'CustomizedPageStatus',
  'ETag',
  'Exists',
  'Length',
  'Level',
  'MajorVersion',
  'MinorVersion',
  'Name',
  'ServerRelativeUrl',
  'TimeCreated',
  'TimeLastModified',
  'Title',
  'UIVersion',
  'UIVersionLabel'
];

const ITEM_PROPS = [
  'AppAuthor',
  'AppEditor',
  'Author',
  'CheckedOutTitle',
  'CheckedOutUserId',
  'CheckoutUser',
  'ContentTypeId',
  'Created',
  'Created_x0020_By',
  'Created_x0020_Date',
  'DocConcurrencyNumber',
  'Editor',
  'FSObjType',
  'FileDirRef',
  'FileLeafRef',
  'FileRef',
  'File_x0020_Size',
  'File_x0020_Type',
  'FolderChildCount',
  'GUID',
  'HTML_x0020_File_x0020_Type',
  'ID',
  'InstanceID',
  'IsCheckedoutToLocal',
  'ItemChildCount',
  'Last_x0020_Modified',
  'MetaInfo',
  'Modified',
  'Modified_x0020_By',
  'Order',
  'ParentLeafName',
  'ParentVersionString',
  'ProgId',
  'ScopeId',
  'SortBehavior',
  'SyncClientId',
  'TemplateUrl',
  'Title',
  'UniqueId',
  'VirusStatus',
  'WorkflowInstanceID',
  'WorkflowVersion',
  'owshiddenversion',
  'xd_ProgID',
  'xd_Signature',
  '_CheckinComment',
  '_CopySource',
  '_HasCopyDestinations',
  '_IsCurrentVersion',
  '_Level',
  '_ModerationComments',
  '_ModerationStatus',
  '_SharedFileIndex',
  '_SourceUrl',
  '_UIVersion',
  '_UIVersionString'
]

const ITEM_REST_PROPS = [
  'AttachmentFiles',
  'AuthorId',
  'CheckoutUserId',
  'ContentType',
  'ContentTypeId',
  'Created',
  'EditorId',
  'FieldValuesAsHtml',
  'FieldValuesForEdit',
  'File',
  'FileSystemObjectType',
  'FirstUniqueAncestorSecurableObject',
  'Folder',
  'GUID',
  'ID',
  'Id',
  'Modified',
  'OData__CopySource',
  'OData__UIVersionString',
  'ParentList',
  'RoleAssignments',
  'Title',
  '__metadata'
]

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const assertObjectItemProps = assertObject(ITEM_PROPS);
const assertCollectionItemProps = assertCollection(ITEM_PROPS);

const assertObjectItemRESTProps = assertObject(ITEM_REST_PROPS);
const assertCollectionItemRESTProps = assertCollection(ITEM_REST_PROPS);

const rootWebList = site().list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2');
const workingWebList = site('test/spx').list('Files');

const singleUrl = 'a/single.txt';
const singleAsStringUrl = 'a/singleAsString.txt';
const singleAsBlobgUrl = 'a/singleAsBlob.txt';
const singleAsItemUrl = 'a/singleAsItem.txt';
const multiUrl = 'a/multiUrl.txt';
const multiUrlAnother = 'a/multiUrlAnother.txt';

const crud = async _ => workingWebList.file(singleUrl).delete({ noRecycle: true, silent: true }).catch(async err => {
  await workingWebList.file(singleUrl).create({ silent: true });
  const newFile = await workingWebList.file(singleUrl).get({ silent: true });
  assert(`Name is not a "single.txt"`)(newFile.Name === 'single.txt');
  const updatedFile = await assertObjectProps('updated file')(workingWebList.file({ Url: singleUrl, Columns: { Title: 'updated file' } }).update({ silent: true }));
  assert(`Title is not a "updated file"`)(updatedFile.Title === 'updated file');
  await workingWebList.file(singleUrl).delete({ noRecycle: true, silent: true });
});

const crudAsSting = async _ => workingWebList.file(singleAsStringUrl).delete({ noRecycle: true, silent: true }).catch(async err => {
  await workingWebList.file({ Url: singleAsStringUrl, Content: 'hi' }).create({ silent: true });
  const newFile = await workingWebList.file(singleAsStringUrl).get({ silent: true, asBlob: true });
  const content = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.srcElement.result);
    reader.readAsText(newFile);
  })
  assert(`content is not a "hi"`)(content === 'hi');
  const updatedFile = await assertObjectProps('updated file')(workingWebList.file({
    Url: singleAsStringUrl,
    Columns: {
      Title: 'updated file'
    },
    Content: 'hi another'
  }).update({ silent: true }));
  assert(`Title is not a "updated file"`)(updatedFile.Title === 'updated file');
  const newUpdatedFile = await workingWebList.file(singleAsStringUrl).get({ silent: true, asBlob: true });
  const contentUpdated = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.srcElement.result);
    reader.readAsText(newUpdatedFile);
  })
  assert(`content is not a "hi another"`)(contentUpdated === 'hi another');
  await workingWebList.file(singleAsStringUrl).delete({ noRecycle: true, silent: true });
});

const crudAsBlob = async _ => workingWebList.file(singleAsStringUrl).delete({ noRecycle: true, silent: true }).catch(async err => {
  await workingWebList.file({ Url: singleAsBlobgUrl, Content: new Blob(['hi'], { type: 'text/csv' }) }).create({ silent: true });
  const newFile = await workingWebList.file(singleAsBlobgUrl).get({ silent: true, asBlob: true });
  const content = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.srcElement.result);
    reader.readAsText(newFile);
  })
  assert(`content is not a "hi"`)(content === 'hi');
  const updatedFile = await assertObjectProps('new file')(workingWebList.file({
    Url: singleAsBlobgUrl,
    Columns: {
      Title: 'updated file'
    },
    Content: 'hi another'
  }).update({ silent: true }));
  assert(`Title is not a "updated file"`)(updatedFile.Title === 'updated file');
  const newUpdatedFile = await workingWebList.file(singleAsBlobgUrl).get({ silent: true, asBlob: true });
  const contentUpdated = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.srcElement.result);
    reader.readAsText(newUpdatedFile);
  })
  assert(`content is not a "hi another"`)(contentUpdated === 'hi another');
  await workingWebList.file(singleAsBlobgUrl).delete({ noRecycle: true, silent: true });
});

const crudAsItem = async _ => workingWebList.file(singleAsItemUrl).delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFile = await assertObjectItemRESTProps('new file')(workingWebList.file({ Url: singleAsItemUrl, Columns: { Title: 'new file' } }).create({ silent: true, asItem: true }));
  assert(`Title is not a "new file"`)(newFile.Title === 'new file');
  const updatedFile = await assertObjectItemProps('updated file')(workingWebList.file({ Url: singleAsItemUrl, Columns: { Title: 'updated file' } }).update({ silent: true, asItem: true }));
  assert(`Title is not a "updated file"`)(updatedFile.Title === 'updated file');
  await workingWebList.file(singleAsItemUrl).delete({ noRecycle: true, silent: true });
});

const crudCollection = async _ => workingWebList.file([multiUrl, multiUrlAnother]).delete({ noRecycle: true, silent: true }).catch(async err => {
  await workingWebList.file([multiUrl, multiUrlAnother]).create({ silent: true });
  const updatedFiles = await assertCollectionProps('updated file')(workingWebList.file([
    { Url: multiUrl, Columns: { Title: 'updated file' } },
    { Url: multiUrlAnother, Columns: { Title: 'updated file another' } }
  ]).update({ silent: true }));
  assert(`Name is not a "multiUrl.txt"`)(updatedFiles[0].Name === 'multiUrl.txt');
  assert(`Name is not a "multiUrlAnother.txt"`)(updatedFiles[1].Name === 'multiUrlAnother.txt');
  await workingWebList.file([multiUrl, multiUrlAnother]).delete({ noRecycle: true, silent: true });
});

const crudCollectionAsItem = async _ => workingWebList.file([multiUrl, multiUrlAnother]).delete({ noRecycle: true, silent: true }).catch(async err => {
  await workingWebList.file([
    { Url: multiUrl, Columns: { Title: 'new file' } },
    { Url: multiUrlAnother, Columns: { Title: 'new file another' } }
  ]).create({ silent: true });
  const newFiles = await workingWebList.file([multiUrl, multiUrlAnother]).get({ silent: true, asItem: true });
  assert(`Title is not a "new file"`)(newFiles[0].Title === 'new file');
  assert(`Title is not a "new file another"`)(newFiles[1].Title === 'new file another');
  const updatedFiles = await assertCollectionItemProps('updated file')(workingWebList.file([
    { Url: multiUrl, Columns: { Title: 'updated file' } },
    { Url: multiUrlAnother, Columns: { Title: 'updated file another' } }
  ]).update({ silent: true, asItem: true }));
  assert(`Title is not a "updated file"`)(updatedFiles[0].Title === 'updated file');
  assert(`Title is not a "updated file another"`)(updatedFiles[1].Title === 'updated file another');
  await workingWebList.file([multiUrl, multiUrlAnother]).delete({ noRecycle: true, silent: true });
});

export default async _ => {
  await Promise.all([
    assertObjectProps('root web list file')(rootWebList.file('simple.aspx').get()),
    assertCollectionProps('root web list file')(rootWebList.file('/').get()),
    assertObjectProps('web list file')(workingWebList.file('test.txt').get()),
    assertCollectionProps('web list root file')(workingWebList.file('/').get()),
    assertCollectionProps('web list test.txt, test.js file')(workingWebList.file(['test.txt', 'test.js']).get()),
    assertObjectProps('web a list folder file')(workingWebList.file('a/add.png').get()),
    assertCollectionProps('web a list folder file')(workingWebList.file('a/').get()),
    assertObjectItemProps('web list file')(workingWebList.file('test.txt').get({ asItem: true })),
    assertCollectionItemProps('web list root file')(workingWebList.file('/').get({ asItem: true })),
    assertCollectionItemProps('web list test.txt, test.js file')(workingWebList.file(['test.txt', 'test.js']).get({ asItem: true })),
    assertObjectItemProps('web a list folder file')(workingWebList.file('a/add.png').get({ asItem: true })),
    assertCollectionItemProps('web a list folder file')(workingWebList.file('a/').get({ asItem: true }))
  ]);
  await crud();
  await crudAsSting();
  await crudAsBlob();
  await crudAsItem();
  await crudCollection();
  await crudCollectionAsItem()
  testIsOk('fileList')()
} 
