/* eslint no-unused-vars:0 */
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
	const keywordName = 'Sharepoint rocks'
	const newKeywordName = 'Sharepoint rocks!!!'
	await web()
		.keyword(keywordName)
		.delete({ silentErrors: true })
		.catch(identity)
	const newKeyword = await assertObjectProps('new keyword')(web()
		.keyword({ Label: keywordName })
		.create({ detailed: true }))
	assert(`Name is not a "${keywordName}"`)(newKeyword.Name === keywordName)
	const updatedkeyword = await assertObjectProps('updated keyword')(web().keyword({
		Label: keywordName,
		Name: newKeywordName
	}).update({ detailed: true }))
	assert(`Name is not a "${newKeywordName}"`)(updatedkeyword.Name === newKeywordName)
	await web().keyword(newKeywordName).delete({ detailed: true })
}

const crudCollection = async () => {
	const keywordName = 'Sharepoint rocks multi'
	const keywordNameAnother = 'Sharepoint rocks another multi'
	const newKeywordName = 'Sharepoint rocks!!! multi'
	const newKeywordNameAnother = 'Sharepoint rocks another!!! multi'
	await web()
		.keyword([keywordName, keywordNameAnother])
		.delete({ silentErrors: true })
		.catch(identity)
	const newkeywords = await assertCollectionProps('new keywords')(
		web().keyword([{ Label: keywordName }, { Label: keywordNameAnother }]).create()
	)
	assert(`Name is not a "${keywordName}"`)(newkeywords[0].Name === keywordName)
	assert(`Name is not a "${keywordNameAnother}"`)(newkeywords[1].Name === keywordNameAnother)
	const updatedkeywords = await assertCollectionProps('updated keyword')(
		web()
			.keyword([{
				Label: keywordName,
				Name: newKeywordName
			}, {
				Label: keywordNameAnother,
				Name: newKeywordNameAnother
			}])
			.update()
			.catch(() => web().keyword([newKeywordName, newKeywordNameAnother]).delete())
	)
	assert(`Name is not a "${newKeywordName}"`)(updatedkeywords[0].Name === newKeywordName)
	assert(`Name is not a "${newKeywordNameAnother}"`)(updatedkeywords[1].Name === newKeywordNameAnother)
	await web().keyword([newKeywordName, newKeywordNameAnother]).delete()
}

export default async () => Promise.all([
	assertObjectProps('keyword "test"')(web().keyword('test').get()),
	assert('keywords are present')((await web().keyword(['test', 'b']).get()).length === 1),
	assert('keywords are missed')((await web().keyword(['test', 'a']).get()).length === 2),
	// crud(),
	// crudCollection()
]).then(testIsOk('keyword'))
