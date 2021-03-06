/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testWrapper,
	assert,
	identity
} from '../lib/utility'

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
]

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

const REST_PROPS = [
	'Author',
	'CheckInComment',
	'CheckOutType',
	'CheckedOutByUser',
	'ContentTag',
	'CustomizedPageStatus',
	'ETag',
	'Exists',
	'ListItemAllFields',
	'LockedByUser',
	'MajorVersion',
	'MinorVersion',
	'ModifiedBy',
	'Name',
	'ServerRelativeUrl',
	'TimeCreated',
	'TimeLastModified',
	'Title',
	'__metadata'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const assertObjectItemProps = assertObject(ITEM_PROPS)
const assertCollectionItemProps = assertCollection(ITEM_PROPS)

const rootWebList = web().library('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2')
const workingWebList = web('test/spx').library('Files')

const crud = async () => {
	const folder = 'a'
	const filename = 'single.txt'
	const url = `${folder}/${filename}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true })
		.catch(identity)
	await workingWebList.file(url).create({ detailed: true })
	const newFile = await workingWebList.file(url).get()
	assert(`Name is not a "${filename}"`)(newFile.Name === filename)
	const updatedFile = await assertObjectProps('updated file')(
		workingWebList.file({ Url: url, Columns: { Title: 'updated file' } }).update()
	)
	assert('Title is not a "updated file"')(updatedFile.Title === 'updated file')
	await workingWebList.file(url).delete({ noRecycle: true })
}

const crudAsSting = async () => {
	const folder = '/test/spx/Files/b'
	const filename = 'singleAsString.txt'
	const url = `${folder}/${filename}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true })
		.catch(identity)
	const createdFile = await workingWebList.file({ Url: url, Content: 'hi', Columns: { Title: 'new file' } }).create()
	assert('Title is not a "new file"')(createdFile.Title === 'new file')
	const newFile = await workingWebList.file(url).get({ asBlob: true })
	const content = await new Promise(resolve => {
		const reader = new FileReader()
		reader.onload = e => resolve(e.srcElement.result)
		reader.readAsText(newFile)
	})
	assert('content is not a "hi"')(content === 'hi')
	const updatedFile = await assertObjectProps('updated file')(
		workingWebList
			.file({
				Url: url,
				Columns: {
					Title: 'updated file'
				},
				Content: 'hi another'
			})
			.update()
	)
	assert('Title is not a "updated file"')(updatedFile.Title === 'updated file')
	const newUpdatedFile = await workingWebList.file(url).get({ asBlob: true })
	const contentUpdated = await new Promise(resolve => {
		const reader = new FileReader()
		reader.onload = e => resolve(e.srcElement.result)
		reader.readAsText(newUpdatedFile)
	})
	assert('content is not a "hi another"')(contentUpdated === 'hi another')
	await workingWebList.file(url).delete({ noRecycle: true })
}

const crudAsBlob = async () => {
	const folder = 'c'
	const filename = 'singleAsBlob.txt'
	const url = `${folder}/${filename}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true, silentErrors: true })
		.catch(identity)
	const createdFile = await workingWebList
		.file({ Url: url, Content: new Blob(['hi'], { type: 'text/csv' }) })
		.create({ needResponse: true })
	// console.log(createdFile);
	const newFile = await workingWebList.file(url).get({ asBlob: true })
	// console.log(newFile);
	const content = await new Promise(resolve => {
		const reader = new FileReader()
		reader.onload = e => resolve(e.srcElement.result)
		reader.readAsText(newFile)
	})
	assert('content is not a "hi"')(content === 'hi')
	const updatedFile = await assertObjectProps('new file')(
		workingWebList
			.file({
				Url: url,
				Columns: {
					Title: 'updated file'
				},
				Content: 'hi another'
			})
			.update()
	)
	assert('Title is not a "updated file"')(updatedFile.Title === 'updated file')
	const newUpdatedFile = await workingWebList.file(url).get({ asBlob: true })
	const contentUpdated = await new Promise(resolve => {
		const reader = new FileReader()
		reader.onload = e => resolve(e.srcElement.result)
		reader.readAsText(newUpdatedFile)
	})
	assert('content is not a "hi another"')(contentUpdated === 'hi another')
	await workingWebList.file(url).delete({ noRecycle: true })
}

const crudAsItem = async () => {
	const folder = 'd'
	const filename = 'singleAsItem.txt'
	const url = `${folder}/${filename}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true, silentErrors: true })
		.catch(identity)
	const newFile = await assertObjectItemProps('new file')(
		workingWebList.file({ Url: url, Columns: { Title: 'new file' } }).create({ asItem: true })
	)
	assert('Title is not a "new file"')(newFile.Title === 'new file')
	const updatedFile = await assertObjectItemProps('updated file')(
		workingWebList.file({ Url: url, Columns: { Title: 'updated file' } }).update({ asItem: true })
	)
	assert('Title is not a "updated file"')(updatedFile.Title === 'updated file')
	await workingWebList.file(url).delete({ noRecycle: true })
}

const crudCollection = async () => {
	const folder = 'e'
	const filename = 'multiUrl.txt'
	const filenameAnother = 'multiUrlAnother.txt'
	const url = `${folder}/${filename}`
	const urlAnother = `${folder}/${filenameAnother}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true, silentErrors: true })
		.catch(identity)
	await workingWebList.file([url, urlAnother]).create()
	const updatedFiles = await assertCollectionProps('updated file')(
		workingWebList
			.file([
				{ Url: url, Columns: { Title: 'updated file' } },
				{ Url: urlAnother, Columns: { Title: 'updated file another' } }
			])
			.update()
	)
	assert(`Name is not a "${filename}"`)(updatedFiles[0].Name === filename)
	assert(`Name is not a "${filenameAnother}"`)(updatedFiles[1].Name === filenameAnother)
	await workingWebList.file([url, urlAnother]).delete({ noRecycle: true })
}

const crudCollectionAsItem = async () => {
	const folder = 'f'
	const filename = 'multiUrl.txt'
	const filenameAnother = 'multiUrlAnother.txt'
	const url = `${folder}/${filename}`
	const urlAnother = `${folder}/${filenameAnother}`
	await workingWebList
		.folder(folder)
		.delete({ noRecycle: true, silentErrors: true })
		.catch(identity)
	await workingWebList
		.file([
			{ Url: url, Columns: { Title: 'new file' } },
			{ Url: urlAnother, Columns: { Title: 'new file another' } }
		])
		.create()
	const newFiles = await workingWebList.file([url, urlAnother]).get({ asItem: true })
	assert('Title is not a "new file"')(newFiles[0].Title === 'new file')
	assert('Title is not a "new file another"')(newFiles[1].Title === 'new file another')
	const updatedFiles = await assertCollectionItemProps('updated file')(
		workingWebList
			.file([
				{ Url: url, Columns: { Title: 'updated file' } },
				{ Url: urlAnother, Columns: { Title: 'updated file another' } }
			])
			.update({ asItem: true })
	)
	assert('Title is not a "updated file"')(updatedFiles[0].Title === 'updated file')
	assert('Title is not a "updated file another"')(updatedFiles[1].Title === 'updated file another')
	await workingWebList.file([url, urlAnother]).delete({ noRecycle: true })
}

export default {
	get: () => testWrapper('list file GET')([
		() => assertObjectProps('root web list file')(rootWebList.file('simple.aspx').get()),
		() => assertCollectionProps('root web list file')(rootWebList.file('/').get()),
		() => assertObjectProps('web list file')(workingWebList.file('test.txt').get()),
		() => assertCollectionProps('web list root file')(workingWebList.file('/').get()),
		() => assertCollectionProps('web list test.txt, test.js file')(workingWebList
			.file(['test.txt', 'test.js'])
			.get()),
		() => assertObjectProps('web a list folder file')(workingWebList.file('g/full.jpg').get()),
		() => assertCollectionProps('web a list folder file')(workingWebList.file('g/').get()),
		() => assertObjectItemProps('web list file')(workingWebList.file('test.txt').get({ asItem: true })),
		() => assertCollectionItemProps('web list root file')(workingWebList.file('/').get({ asItem: true })),
		() => assertCollectionItemProps('web list test.txt, test.js file')(workingWebList
			.file(['test.txt', 'test.js'])
			.get({ asItem: true })),
		() => assertObjectItemProps('web a list folder file')(workingWebList.file('g/full.jpg').get({ asItem: true })),
		() => assertCollectionItemProps('web a list folder file')(workingWebList.file('g/').get({ asItem: true }))
	]),
	crud: () => testWrapper('list file CRUD')([crud]),
	crudCollection: () => testWrapper('list file CRUD Collection')([crudCollection]),
	crudAsItem: () => testWrapper('list file CRUD asItem')([crudAsItem]),
	crudCollectionAsItem: () => testWrapper('list file CRUD Collection as item')([crudCollectionAsItem]),
	crudAsSting: () => testWrapper('list file CRUD as string')([crudAsSting]),
	crudAsBlob: () => testWrapper('list file CRUD Collection as blob')([crudAsBlob]),
	all() {
		testWrapper('list file ALL')([
			this.get,
			this.crud,
			this.crudCollection,
			this.crudAsItem,
			this.crudCollectionAsItem,
			this.crudAsSting,
			this.crudAsBlob
		])
	}
}
