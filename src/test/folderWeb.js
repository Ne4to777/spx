import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, identity } from './../lib/utility';

const PROPS = [
  'ItemCount',
  'Name',
  'ServerRelativeUrl',
  'WelcomePage'
];

const assertObjectProps = assertObject(PROPS);
const assertCollectionProps = assertCollection(PROPS);

const rootWeb = site();
const workingWeb = site('test/spx');

const crud = async _ => {
  await workingWeb.folder('singleFolder').delete({ noRecycle: true, silent: true }).catch(identity);
  const newFolder = await assertObjectProps('new folder')(workingWeb.folder('singleFolder').create());
  assert(`Name is not a "singleFolder"`)(newFolder.Name === 'singleFolder');
  await workingWeb.folder('singleFolder').delete({ noRecycle: true });
}

const crudCollection = async _ => {
  await workingWeb.folder(['multiFolder', 'multiFolderAnother']).delete({ noRecycle: true, silent: true }).catch(identity);
  const newFolders = await assertCollectionProps('new folder')(workingWeb.folder(['multiFolder', 'multiFolderAnother']).create());
  assert(`Name is not a "multiFolder"`)(newFolders[0].Name === 'multiFolder');
  assert(`Name is not a "multiFolderAnother"`)(newFolders[1].Name === 'multiFolderAnother');
  await workingWeb.folder(['multiFolder', 'multiFolderAnother']).delete({ noRecycle: true });
}

const crudBundle = async _ => {
  const foldersToCreate = [];
  const folder = 'b'
  for (let i = 0; i < 253; i++) foldersToCreate.push(`${folder}/test${i}`);
  // console.log(foldersToCreate);
  await workingWeb.folder(foldersToCreate).delete({ noRecycle: true }).catch(identity);
  await workingWeb.folder(foldersToCreate).create();
  await workingWeb.folder(foldersToCreate).delete({ noRecycle: true });
}

export default _ => Promise.all([
  assertObjectProps('root web folder')(rootWeb.folder().get()),
  assertCollectionProps('root web folder')(rootWeb.folder('/').get()),
  assertObjectProps('web root folder')(workingWeb.folder().get()),
  assertObjectProps('web _catalogs folder')(workingWeb.folder('_catalogs').get()),
  assertCollectionProps('web _catalogs folder')(workingWeb.folder('_catalogs/').get()),
  assertCollectionProps('web _catalogs folder')(workingWeb.folder(['_catalogs', 'Files']).get()),
  crud(),
  crudCollection(),
  // crudBundle()
]).then(testIsOk('folderWeb'))