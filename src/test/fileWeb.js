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

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = web()
const workingWeb = web('test/spx')

const crud = async () => {
	const folder = 'b'
	const filename = 'single.txt'
	const url = `${folder}/${filename}`
	await workingWeb
		.folder(folder)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFile = await assertObjectProps('new file')(workingWeb.file({ Url: url }).create())
	assert('Name is not a "single.txt"')(newFile.Name === 'single.txt')
	await workingWeb.file({ Url: url, Content: 'updated' }).update()
	await workingWeb.file({ Url: url }).delete({ noRecycle: true })
}

const crudCollection = async () => {
	const folder = 'b'
	const filename = 'multi.txt'
	const filenameAnother = 'multiAnother.txt'
	const url = `${folder}/${filename}`
	const urlAnother = `${folder}/${filenameAnother}`
	await workingWeb
		.folder(folder)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFiles = await assertCollectionProps('new file')(workingWeb.file([url, urlAnother]).create())
	assert(`Name is not a "${filename}"`)(newFiles[0].Name === filename)
	assert(`Name is not a "${filenameAnother}"`)(newFiles[1].Name === filenameAnother)
	await workingWeb.file([url, urlAnother]).delete({ noRecycle: true })
}

const crudBundle = async () => {
	const foldersToCreate = []
	const folder = 'b'
	for (let i = 0; i < 253; i += 1) foldersToCreate.push(`${folder}/single${i}.txt`)
	// console.log(foldersToCreate);
	await workingWeb
		.folder(folder)
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	await workingWeb.folder(foldersToCreate).create()
	await workingWeb.folder(foldersToCreate).delete({ noRecycle: true })
}

export default {
	get: () => testWrapper('web file GET')([
		() => assertObjectProps('root web file')(rootWeb.file('index.html').get()),
		() => assertCollectionProps('root web file')(rootWeb.file('/').get()),
		() => assertObjectProps('web file')(workingWeb.file('index.aspx').get()),
		() => assertCollectionProps('web root file')(workingWeb.file('/').get()),
		() => assertCollectionProps('web index.aspx, default.aspx file')(workingWeb
			.file(['index.aspx', 'default.aspx'])
			.get()),
	]),
	crud: () => testWrapper('web file CRUD')([crud]),
	crudCollection: () => testWrapper('web file CRUD Collection')([crudCollection]),
	crudBundle: () => testWrapper('web file CRUD Bundle')([crudBundle]),
	all: () => testWrapper('web file ALL')([
		this.get,
		this.crud,
		this.crudCollection,
		this.crudBundle
	])
}
