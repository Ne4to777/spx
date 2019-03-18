import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert } from './../lib/utility';

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

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const rootWeb = site();
const workingWeb = site('test/spx');

const crud = async _ => {
  await workingWeb.file({ Url: 'a/single.txt' }).delete({ noRecycle: true, silent: true })
  const newFile = await assertObjectProps('new file')(workingWeb.file({ Url: 'a/single.txt' }).create({ silent: true }));
  assert(`Name is not a "single.txt"`)(newFile.Name === 'single.txt');
  await workingWeb.file({ Url: 'a/single.txt' }).delete({ noRecycle: true, silent: true });
};

const crudCollection = async _ => {
  await workingWeb.file(['multi.txt', 'multiAnother.txt']).delete({ noRecycle: true, silent: true })
  const newFiles = await assertCollectionProps('new file')(workingWeb.file(['multi.txt', 'multiAnother.txt']).create({ silent: true }));
  assert(`Name is not a "multi.txt"`)(newFiles[0].Name === 'multi.txt');
  assert(`Name is not a "multiAnother.txt"`)(newFiles[1].Name === 'multiAnother.txt');
  await workingWeb.file(['multi.txt', 'multiAnother.txt']).delete({ noRecycle: true, silent: true });
};

export default _ => Promise.all([
  assertObjectProps('root web file')(rootWeb.file('index.html').get()),
  assertCollectionProps('root web file')(rootWeb.file('/').get()),
  assertObjectProps('web file')(workingWeb.file('index.aspx').get()),
  assertCollectionProps('web root file')(workingWeb.file('/').get()),
  assertCollectionProps('web index.aspx, default.aspx file')(workingWeb.file(['index.aspx', 'default.aspx']).get()),
  crud(),
  crudCollection()
]).then(testIsOk('fileWeb'))