import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert } from './../lib/utility';

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

const crud = async _ => workingWeb.folder('singleFolder').delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolder = await assertObjectProps('new folder')(workingWeb.folder('singleFolder').create({ silent: true }));
  assert(`Name is not a "singleFolder"`)(newFolder.Name === 'singleFolder');
  await workingWeb.folder('singleFolder').delete({ noRecycle: true, silent: true });
});

const crudCollection = async _ => workingWeb.folder(['multiFolder', 'multiFolderAnother']).delete({ noRecycle: true, silent: true }).catch(async err => {
  const newFolders = await assertCollectionProps('new folder')(workingWeb.folder(['multiFolder', 'multiFolderAnother']).create({ silent: true }));
  assert(`Name is not a "multiFolder"`)(newFolders[0].Name === 'multiFolder');
  assert(`Name is not a "multiFolderAnother"`)(newFolders[1].Name === 'multiFolderAnother');
  await workingWeb.folder(['multiFolder', 'multiFolderAnother']).delete({ noRecycle: true, silent: true });
});

export default _ => Promise.all([
  assertObjectProps('root web folder')(rootWeb.folder().get()),
  assertCollectionProps('root web folder')(rootWeb.folder('/').get()),
  assertObjectProps('web root folder')(workingWeb.folder().get()),
  assertObjectProps('web _catalogs folder')(workingWeb.folder('_catalogs').get()),
  assertCollectionProps('web _catalogs folder')(workingWeb.folder('_catalogs/').get()),
  assertCollectionProps('web _catalogs folder')(workingWeb.folder(['_catalogs', 'Files']).get()),
  crud(),
  crudCollection()
]).then(testIsOk('folderWeb'))