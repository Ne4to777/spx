import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testIsOk,
	assert,
	identity
} from '../lib/utility'

const PROPS = [
	'CreatedDate',
	'CustomProperties',
	'CustomSortOrder',
	'Description',
	'Id',
	'IsAvailableForTagging',
	'IsDeprecated',
	'IsKeyword',
	'IsPinned',
	'IsPinnedRoot',
	'IsReused',
	'IsRoot',
	'IsSourceTerm',
	'LastModifiedDate',
	'LocalCustomProperties',
	'MergedTermIds',
	'Name',
	'Owner',
	'PathOfTerm',
	'TermsCount'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const crud = async () => {
	const tagName = 'Sharepoint rocks'
	const newTagName = 'Sharepoint rocks!!!'
	await web
		.tag(tagName)
		.delete({ silentErrors: true })
		.catch(identity)
	const newTag = await assertObjectProps('new tag')(web.tag({ Url: tagName }).create())
	assert(`Name is not a "${tagName}"`)(newTag.Name === tagName)
	const updatedTag = await assertObjectProps('updated tag')(web.tag({ Url: tagName, Name: newTagName }).update())
	assert(`Name is not a "${newTagName}"`)(updatedTag.Name === newTagName)
	await web.tag(newTagName).delete()
}

const crudCollection = async () => {
	const tagName = 'Sharepoint rocks multi'
	const tagNameAnother = 'Sharepoint rocks another multi'
	const newTagName = 'Sharepoint rocks!!! multi'
	const newTagNameAnother = 'Sharepoint rocks another!!! multi'
	await web
		.tag([tagName, tagNameAnother])
		.delete({ silentErrors: true })
		.catch(identity)
	const newTags = await assertCollectionProps('new tags')(
		web.tag([{ Url: tagName }, { Url: tagNameAnother }]).create()
	)
	assert(`Name is not a "${tagName}"`)(newTags[0].Name === tagName)
	assert(`Name is not a "${tagNameAnother}"`)(newTags[1].Name === tagNameAnother)
	const updatedTags = await assertCollectionProps('updated tag')(
		web
			.tag([{ Url: tagName, Name: newTagName }, { Url: tagNameAnother, Name: newTagNameAnother }])
			.update()
			.catch(() => web.tag([newTagName, newTagNameAnother]).delete())
	)
	assert(`Name is not a "${newTagName}"`)(updatedTags[0].Name === newTagName)
	assert(`Name is not a "${newTagNameAnother}"`)(updatedTags[1].Name === newTagNameAnother)
	await web.tag([newTagName, newTagNameAnother]).delete()
}

export default async () => Promise.all([
	assertObjectProps('tag "test"')(web.tag('test').get()),
	assert('tags are present')((await web.tag(['test', 'b']).get()).length === 1),
	assert('tags are missed')((await web.tag(['test', 'a']).get()).length === 2),
	crud(),
	crudCollection()
]).then(testIsOk('tag'))
