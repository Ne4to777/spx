import site from './../modules/site'
import { assertObject, assertCollection, testIsOk, assert, identity } from './../lib/utility'

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

const crud = async _ => {
	const tagName = 'Sharepoint rocks'
	const newTagName = 'Sharepoint rocks!!!'
	await site
		.tag(tagName)
		.delete({ silentErrors: true })
		.catch(identity)
	const newTag = await assertObjectProps('new tag')(site.tag({ Url: tagName }).create())
	assert(`Name is not a "${tagName}"`)(newTag.Name === tagName)
	const updatedTag = await assertObjectProps('updated tag')(site.tag({ Url: tagName, Name: newTagName }).update())
	assert(`Name is not a "${newTagName}"`)(updatedTag.Name === newTagName)
	await site.tag(newTagName).delete()
}

const crudCollection = async _ => {
	const tagName = 'Sharepoint rocks multi'
	const tagNameAnother = 'Sharepoint rocks another multi'
	const newTagName = 'Sharepoint rocks!!! multi'
	const newTagNameAnother = 'Sharepoint rocks another!!! multi'
	await site
		.tag([tagName, tagNameAnother])
		.delete({ silentErrors: true })
		.catch(identity)
	const newTags = await assertCollectionProps('new tags')(
		site.tag([{ Url: tagName }, { Url: tagNameAnother }]).create()
	)
	assert(`Name is not a "${tagName}"`)(newTags[0].Name === tagName)
	assert(`Name is not a "${tagNameAnother}"`)(newTags[1].Name === tagNameAnother)
	const updatedTags = await assertCollectionProps('updated tag')(
		site
			.tag([{ Url: tagName, Name: newTagName }, { Url: tagNameAnother, Name: newTagNameAnother }])
			.update()
			.catch(_ => site.tag([newTagName, newTagNameAnother]).delete())
	)
	assert(`Name is not a "${newTagName}"`)(updatedTags[0].Name === newTagName)
	assert(`Name is not a "${newTagNameAnother}"`)(updatedTags[1].Name === newTagNameAnother)
	await site.tag([newTagName, newTagNameAnother]).delete()
}

export default async _ =>
	Promise.all([
		assertObjectProps('tag "test"')(site.tag('test').get()),
		assert(`tags are present`)((await site.tag(['test', 'b']).get()).length === 1),
		assert(`tags are missed`)((await site.tag(['test', 'a']).get()).length === 2),
		crud(),
		crudCollection()
	]).then(testIsOk('tag'))
