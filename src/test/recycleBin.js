
import web from '../modules/web'
import {
	assertCollection,
	testWrapper
} from '../lib/utility'

const PROPS = [
	'DeletedDate',
	'DirName',
	'Id',
	'ItemState',
	'ItemType',
	'LeafName',
	'Size',
	'Title',
]

const assertPropsCollection = assertCollection(PROPS)

export default {
	get: () => testWrapper('recycleBin GET')([
		() => assertPropsCollection('recycleBin collection')(web('test/spx').recycleBin().get()),
	]),
	all() { testWrapper('time ALL')([this.get]) }
}
