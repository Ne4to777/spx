/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testWrapper,
	assert,
	identity
} from '../lib/utility'

const PROPS = ['ItemCount', 'Name', 'ServerRelativeUrl', 'WelcomePage']

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = web()
const workingWeb = web('test/spx')

const crud = async () => {
	await workingWeb
		.folder('singleFolder')
		.delete({ noRecycle: true, silent: true })
		.catch(identity)

	const newFolder = await assertObjectProps('new folder')(workingWeb.folder('singleFolder').create())
	assert('Name is not a "singleFolder"')(newFolder.Name === 'singleFolder')
	await workingWeb.folder('singleFolder').delete({ noRecycle: true })
}

const crudCollection = async () => {
	await workingWeb
		.folder(['multiFolder', 'multiFolderAnother'])
		.delete({ noRecycle: true, silent: true })
		.catch(identity)
	const newFolders = await assertCollectionProps('new folder')(
		workingWeb.folder(['multiFolder', 'multiFolderAnother']).create()
	)
	assert('Name is not a "multiFolder"')(newFolders[0].Name === 'multiFolder')
	assert('Name is not a "multiFolderAnother"')(newFolders[1].Name === 'multiFolderAnother')
	await workingWeb.folder(['multiFolder', 'multiFolderAnother']).delete({ noRecycle: true })
}

const crudBundle = async () => {
	const foldersToCreate = []
	const folder = 'b'
	for (let i = 0; i < 253; i += 1) foldersToCreate.push(`${folder}/test${i}`)
	// console.log(foldersToCreate)
	await workingWeb
		.folder(foldersToCreate)
		.delete({ noRecycle: true })
		.catch(identity)
	await workingWeb
		.folder(foldersToCreate)
		.create()
	await workingWeb
		.folder(foldersToCreate)
		.delete({ noRecycle: true })
}

export default {
	get: () => testWrapper('web folder GET')([
		() => assertObjectProps('root web folder')(rootWeb.folder().get()),
		() => assertCollectionProps('root web folder')(rootWeb.folder('/').get()),
		() => assertObjectProps('web root folder')(workingWeb.folder().get()),
		() => assertObjectProps('web _catalogs folder')(workingWeb.folder('_catalogs').get()),
		() => assertCollectionProps('web _catalogs folder')(workingWeb.folder('_catalogs/').get()),
		() => assertCollectionProps('web _catalogs folder')(workingWeb.folder(['_catalogs', 'Files']).get()),
	]),
	crud: () => testWrapper('web folder CRUD')([crud]),
	crudCollection: () => testWrapper('web folder CRUD Collection')([crudCollection]),
	crudBundle: () => testWrapper('web folder CRUD Bundle')([crudBundle]),
	all: () => testWrapper('web folder ALL')([
		this.get,
		this.crud,
		this.crudCollection,
		this.crudBundle
	])
}
