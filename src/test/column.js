import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testIsOk,
	assert,
	identity
} from '../lib/utility'

const PROPS = [
	'CanBeDeleted',
	'DefaultValue',
	'Description',
	'Direction',
	'EnforceUniqueValues',
	'EntityPropertyName',
	'FieldTypeKind',
	'Filterable',
	'FromBaseType',
	'Group',
	'Hidden',
	'Id',
	'Indexed',
	'InternalName',
	'JSLink',
	'ReadOnlyField',
	'Required',
	'SchemaXml',
	'Scope',
	'Sealed',
	'Sortable',
	'StaticName',
	'Title',
	'TypeAsString',
	'TypeDisplayName',
	'TypeShortDescription',
	'ValidationFormula',
	'ValidationMessage'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const rootWeb = web()
const workingWeb = web('test/spx')

const crud = async () => {
	await workingWeb
		.list('Test')
		.column('single')
		.delete({ silent: true })
		.catch(identity)
	const newColumn = await assertObjectProps('new column')(
		workingWeb
			.list('Test')
			.column({ Title: 'single', Description: 'new column' })
			.create()
	)
	assert('Description is not a "new column"')(newColumn.Description === 'new column')
	const updatedColumn = await assertObjectProps('updated column')(
		workingWeb
			.list('Test')
			.column({ Title: 'single', Description: 'updated column' })
			.update()
	)
	assert('Description is not a "updated column"')(updatedColumn.Description === 'updated column')
	await workingWeb
		.list('Test')
		.column('single')
		.delete()
}

const crudCollection = async () => {
	await workingWeb
		.list('Test')
		.column(['multi', 'multiAnother'])
		.delete({ silent: true })
		.catch(identity)
	const newColumns = await assertCollectionProps('new column')(
		workingWeb
			.list('Test')
			.column([
				{
					Title: 'multi',
					Description: 'new multi column'
				},
				{
					Title: 'multiAnother',
					Description: 'new multiAnother column'
				}
			])
			.create()
	)
	assert('Description is not a "new multi column"')(newColumns[0].Description === 'new multi column')
	assert('Description is not a "new multiAnother column"')(newColumns[1].Description === 'new multiAnother column')
	const updatedColumns = await assertCollectionProps('updated column')(
		workingWeb
			.list('Test')
			.column([
				{
					Title: 'multi',
					Description: 'updated multi column'
				},
				{
					Title: 'multiAnother',
					Description: 'updated multiAnother column'
				}
			])
			.update()
	)
	assert('Description is not a "updated multi column"')(updatedColumns[0].Description === 'updated multi column')
	assert('Description is not a "updated multiAnother column"')(
		updatedColumns[1].Description === 'updated multiAnother column'
	)
	await workingWeb
		.list('Test')
		.column([
			{
				Title: 'multi'
			},
			{
				Title: 'multiAnother'
			}
		])
		.delete()
}

export default () => Promise.all([
	assertObjectProps('root web list column')(
		rootWeb
			.list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2')
			.column('Title')
			.get()
	),
	assertCollectionProps('root web list column')(
		rootWeb
			.list('b327d30a-b9bf-4728-a3c1-a6b4f0253ff2')
			.column()
			.get()
	),
	assertObjectProps('web list column')(
		workingWeb
			.list('Test')
			.column('Title')
			.get()
	),
	assertCollectionProps('web root list column')(
		workingWeb
			.list('Test')
			.column()
			.get()
	),
	assertCollectionProps('web Test, TestAnother list column')(
		workingWeb
			.list(['Test', 'TestAnother'])
			.column(['Title', 'Author'])
			.get()
	),
	crud().then(crudCollection)
]).then(testIsOk('column'))
