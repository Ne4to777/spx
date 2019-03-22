import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, map, prop, reduce, identity } from './../lib/utility';

const PROPS = [
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
  'Title',
  'UniqueId',
  'WorkflowInstanceID',
  'WorkflowVersion',
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

const USER_PROPS = [
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

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const assertObjectUserProps = assertObject(USER_PROPS);
const assertCollectionUserProps = assertCollection(USER_PROPS);

const userWebList = site().list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2');
const workingWebList = site('test/spx').list('Items');

const crud = async _ => {
  const folder = 'a';
  await workingWebList.folder(folder).delete({ noRecycle: true }).catch(identity);
  const newItem = await workingWebList.item({ Folder: folder, Columns: { Title: 'new item' } }).create({ view: ['ID', 'Title', 'FileDirRef'] });
  assert(`Title is not a "new item"`)(newItem.Title === 'new item');
  assert(`FileDirRef is not a "/test/spx/Lists/Items/a"`)(newItem.FileDirRef === '/test/spx/Lists/Items/a');
  const updatedItem = await workingWebList.item({ ID: newItem.ID, Title: 'updated item' }).update();
  assert(`Title is not a "updated item"`)(updatedItem.Title === 'updated item');
  await workingWebList.item(newItem.ID).delete({ noRecycle: true });
};


const crudCollection = async _ => {
  const folder = 'b';
  const newItems = await workingWebList.item([
    {
      Folder: folder,
      Columns: {
        Title: 'new item'
      }
    }, {
      Folder: folder,
      Columns: {
        Title: 'new item another'
      }
    }]).create({ view: ['FileDirRef', 'ID', 'Title'] });
  // console.log(newItems);
  assert(`Title is not a "new item"`)(newItems[0].Title === 'new item');
  assert(`Title is not a "new item another"`)(newItems[1].Title === 'new item another');
  assert(`FileDirRef is not a "/test/spx/Lists/Items/${folder}"`)(newItems[0].FileDirRef === `/test/spx/Lists/Items/${folder}`);
  assert(`FileDirRef is not a "/test/spx/Lists/Items/${folder}"`)(newItems[1].FileDirRef === `/test/spx/Lists/Items/${folder}`);
  const updatedItems = await workingWebList.item([
    {
      ID: newItems[0].ID,
      Title: 'updated item'
    }, {
      ID: newItems[1].ID,
      Title: 'updated item another'
    }]).update({ view: ['FileDirRef', 'ID'] });
  assert(`Title is not a "updated item"`)(updatedItems[0].Title === 'updated item');
  assert(`Title is not a "updated item another"`)(updatedItems[1].Title === 'updated item another');
  assert(`FileDirRef is not a "/test/spx/Lists/Items/${folder}"`)(updatedItems[0].FileDirRef === `/test/spx/Lists/Items/${folder}`);
  assert(`FileDirRef is not a "/test/spx/Lists/Items/${folder}"`)(updatedItems[1].FileDirRef === `/test/spx/Lists/Items/${folder}`);
  await workingWebList.item([newItems[0].ID, newItems[1].ID]).delete({ noRecycle: true });
};

const crudBundle = async _ => {
  const itemsToCreate = [];
  const budleList = site('test/spx').list('Bundle');
  const folder = 'c/b';
  await budleList.folder(folder).delete({ noRecycle: true }).catch(identity);
  for (let i = 0; i < 253; i++) itemsToCreate.push({ Title: `test ${i}`, Folder: folder })
  const newItems = await budleList.item(itemsToCreate).create();
  const itemsToUpdate = reduce(acc => el => acc.concat({ ID: el.ID, Title: `${el.Title} updated` }))([])(newItems)
  const updatedItems = await budleList.item(itemsToUpdate).update();
  const itemsTodelete = map(prop('ID'))(updatedItems);
  await budleList.item(itemsTodelete).delete({ noRecycle: true });
}



export default _ => Promise.all([
  assertObjectUserProps('user web list item 10842 ID')(userWebList.item(10842).get()),
  assertCollectionUserProps('user web list item first')(userWebList.item({ Limit: 1 }).get()),
  assertCollectionUserProps('user web list item 10842 ID')(userWebList.item('ID eq 10842').get()),
  assertCollectionProps('web list item')(workingWebList.item().get()),
  assertObjectProps('web a list item 305 ID ')(workingWebList.item(305).get()),
  assertCollectionProps('web d list item')(workingWebList.item({ Folder: 'd' }).get()),
  assertCollectionProps('web d, e list item')(workingWebList.item([{ Folder: 'd' }, { Folder: 'e' }]).get()),

  crud(),
  crudCollection(),
  crudBundle()

]).then(testIsOk('itemList'))