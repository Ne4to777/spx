import site from '../modules/site'
import {
	assertObject,
	assertCollection,
	testIsOk,
	assert,
	map,
	prop,
	reduce,
	identity
} from '../lib/utility'

const PROPS = ['ItemCount', 'Name', 'ServerRelativeUrl', 'WelcomePage']

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

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const assertObjectItemProps = assertObject(ITEM_PROPS)
const assertCollectionItemProps = assertCollection(ITEM_PROPS)

const rootWebList = site().list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2')
const workingWebList = site('test/spx').list('Test')

const crud = async () => {
	const folder = 'a/sub'
	const url = `${folder}/singleFolder`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFolder = await assertObjectProps('new folder')(workingWebList.folder(url).create())
	assert('Name is not a "singleFolder"')(newFolder.Name === 'singleFolder')
	const updatedFolder = await assertObjectProps('new folder')(
		workingWebList.folder({ Url: url, Title: 'updated folder' }).update()
	)
	assert('Name is not a "singleFolder"')(updatedFolder.Name === 'singleFolder')
	await workingWebList.folder(url).delete({ noRecycle: true })
}

const crudAsItem = async () => {
	const folder = 'b/sub'
	const url = `${folder}/singleFolderAsItem`
	await workingWebList
		.folder(url)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFolder = await assertObjectItemProps('new folder')(workingWebList.folder(url).create({ asItem: true }))
	assert('FileLeafRef is not a "singleFolderAsItem"')(newFolder.FileLeafRef === 'singleFolderAsItem')
	const updatedFolder = await assertObjectItemProps('new folder')(
		workingWebList.folder({ Url: url, Title: 'updated folder' }).update({ asItem: true })
	)
	assert('Title is not a "updated folder"')(updatedFolder.Title === 'updated folder')
	await workingWebList.folder(url).delete({ noRecycle: true })
}

const crudCollection = async () => {
	const folder = 'c/sub'
	const url = `/test/spx/Lists/Test/${folder}/multiFolder`
	const urlAnother = `/test/spx/Lists/Test/${folder}/multiFolderAnother`
	await workingWebList
		.folder([url, urlAnother])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFolders = await assertCollectionProps('new folder')(workingWebList.folder([url, urlAnother]).create())
	assert('Name is not a "multiFolder"')(newFolders[0].Name === 'multiFolder')
	assert('Name is not a "multiFolderAnother"')(newFolders[1].Name === 'multiFolderAnother')
	const updatedFolder = await assertCollectionProps('new folder')(
		workingWebList
			.folder([{ Url: url, Title: 'updated folder' }, { Url: urlAnother, Title: 'updated folder another' }])
			.update({ silent: false })
	)
	assert('Name is not a "multiFolder"')(updatedFolder[0].Name === 'multiFolder')
	assert('Name is not a "multiFolderAnother"')(updatedFolder[1].Name === 'multiFolderAnother')
	await workingWebList.folder([url, urlAnother]).delete({ noRecycle: true })
}

const crudCollectionAsItem = async () => {
	const folder = 'd/sub'
	const url = `/test/spx/Lists/Test/${folder}/multiFolder`
	const urlAnother = `/test/spx/Lists/Test/${folder}/multiFolderAnother`
	await workingWebList
		.folder([url, urlAnother])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFolders = await assertCollectionItemProps('new folder')(
		workingWebList
			.folder([{ Url: url, Title1: 'new folder' }, { Url: urlAnother, Title1: 'new folder another' }])
			.create({ asItem: true })
	)
	assert('Title1 is not a "new folder"')(newFolders[0].Title1 === 'new folder')
	assert('Title1 is not a "new folder another"')(newFolders[1].Title1 === 'new folder another')
	const updatedFolders = await assertCollectionItemProps('new folder')(
		workingWebList
			.folder([{ Url: url, Title1: 'updated folder' }, { Url: urlAnother, Title1: 'updated folder another' }])
			.update({ asItem: true })
	)
	assert('Title1 is not a "updated folder"')(updatedFolders[0].Title1 === 'updated folder')
	assert('Title1 is not a "updated folder another"')(updatedFolders[1].Title1 === 'updated folder another')
	await workingWebList.folder([url, urlAnother]).delete({ noRecycle: true })
}

const crudBundle = async () => {
	const count = 260
	const foldersToCreate = []
	const folder = 'd/sub'
	const bundleList = site('test/spx').list('Bundle')
	await bundleList
		.folder(folder)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	for (let i = 0; i < count; i += 1) {
		foldersToCreate.push({ ServerRelativeUrl: `${folder}/test${i}`, Name: `test${i}` })
	}
	const newFolders = await bundleList.folder(foldersToCreate).create({ detailed: true })
	console.log(newFolders)
	const foldersToUpdate = reduce(acc => el => acc.concat({ Url: el.ServerRelativeUrl, Title: `${el.Name} updated` }))(
		[]
	)(newFolders)
	console.log(foldersToUpdate)
	const updatedItems = await bundleList.folder(foldersToUpdate).update()
	const foldersToDelete = map(prop('ServerRelativeUrl'))(updatedItems)
	console.log(foldersToDelete)
	await bundleList.folder(foldersToCreate).delete({ noRecycle: true })
}

export default () => Promise.all([
	// assertObjectProps('root web list folder')(rootWebList.folder().get()),
	// assertCollectionProps('root web list folder')(rootWebList.folder('/').get()),
	// assertObjectProps('web root list folder')(workingWebList.folder().get()),
	// assertObjectProps('web a list folder')(workingWebList.folder('a').get()),
	// assertCollectionProps('web a list folder')(workingWebList.folder('a/').get()),
	// assertCollectionProps('web a, c list folder')(workingWebList.folder(['a', 'c']).get()),

	// assertObjectItemProps('web a list folder')(workingWebList.folder('a').get({ asItem: true })),
	// assertCollectionItemProps('web a list folder')(workingWebList.folder('a/').get({ asItem: true })
	// .then(filter(isObjectFilled))),
	// assertCollectionItemProps('web a, c list folder')(workingWebList.folder(['a', 'c']).get({ asItem: true })),

	crud(),
	crudCollection(),
	crudAsItem(),
	crudCollectionAsItem()

	// crudBundle()
]).then(testIsOk('folderList'))
