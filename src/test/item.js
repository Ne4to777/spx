import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, map, prop, reduce } from './../lib/utility';

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
const workingWebList = site('test/spx').list('Test');

const crud = async _ => {
  const newItem = await workingWebList.item({ Folder: 'a', Columns: { Title: 'new item' } }).create({ silent: true, view: ['FileDirRef', 'ID'] });
  assert(`Title is not a "new item"`)(newItem.Title === 'new item');
  assert(`FileDirRef is not a "/test/spx/Lists/Test/a"`)(newItem.FileDirRef === '/test/spx/Lists/Test/a');
  const updatedItem = await workingWebList.item({ ID: newItem.ID, Title: 'updated item' }).update({ silent: true });
  assert(`Title is not a "updated item"`)(updatedItem.Title === 'updated item');
  await workingWebList.item(newItem.ID).delete({ noRecycle: true, silent: true });
};


const crudCollection = async _ => {
  const newItems = await workingWebList.item([
    {
      Folder: 'a',
      Columns: {
        Title: 'new item'
      }
    }, {
      Folder: 'a',
      Columns: {
        Title: 'new item another'
      }
    }]).create({ silent: true, view: ['FileDirRef', 'ID'] });
  assert(`Title is not a "new item"`)(newItems[0].Title === 'new item');
  assert(`Title is not a "new item another"`)(newItems[1].Title === 'new item another');
  assert(`FileDirRef is not a "/test/spx/Lists/Test/a"`)(newItems[0].FileDirRef === '/test/spx/Lists/Test/a');
  assert(`FileDirRef is not a "/test/spx/Lists/Test/a"`)(newItems[1].FileDirRef === '/test/spx/Lists/Test/a');
  const updatedItems = await workingWebList.item([
    {
      ID: newItems[0].ID,
      Title: 'updated item'
    }, {
      ID: newItems[1].ID,
      Title: 'updated item another'
    }]).update({ silent: true, view: ['FileDirRef', 'ID'] });
  assert(`Title is not a "updated item"`)(updatedItems[0].Title === 'updated item');
  assert(`Title is not a "updated item another"`)(updatedItems[1].Title === 'updated item another');
  assert(`FileDirRef is not a "/test/spx/Lists/Test/a"`)(updatedItems[0].FileDirRef === '/test/spx/Lists/Test/a');
  assert(`FileDirRef is not a "/test/spx/Lists/Test/a"`)(updatedItems[1].FileDirRef === '/test/spx/Lists/Test/a');
  await workingWebList.item([newItems[0].ID, newItems[1].ID]).delete({ noRecycle: true, silent: true });
};

const crudBundle = async _ => {
  const itemsToCreate = [];
  const budleList = site('test/spx').list('Bundle');
  for (let i = 0; i < 253; i++) itemsToCreate.push({ Title: `test ${i}` })
  const newItems = await budleList.item(itemsToCreate).create();
  const itemsToUpdate = reduce(acc => el => acc.concat({ ID: el.ID, Title: `${el.Title} updated` }))([])(newItems)
  const updatedItems = await budleList.item(itemsToUpdate).update();
  const itemsTodelete = map(prop('ID'))(updatedItems);
  await budleList.item(itemsTodelete).delete({ noRecycle: true });
}



export default _ => Promise.all([
  // assertObjectUserProps('user web list item 10842 ID')(userWebList.item(10842).get()),
  // assertCollectionUserProps('user web list item first')(userWebList.item({ Limit: 1 }).get()),
  // assertCollectionUserProps('user web list item 10842 ID')(userWebList.item('ID eq 10842').get()),
  // assertCollectionProps('web list item')(workingWebList.item().get()),
  // assertObjectProps('web a list item 1399 ID ')(workingWebList.item(1399).get()),
  // assertCollectionProps('web a list item a folder')(workingWebList.item({ Folder: 'a' }).get()),
  // assertCollectionProps('web a, c list item')(workingWebList.item([{ Folder: 'a' }, { Folder: 'c' }]).get()),

  // crud(),
  // crudCollection(),
  // crudBundle()

]).then(testIsOk('itemList'))