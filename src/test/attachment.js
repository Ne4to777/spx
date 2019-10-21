
import web from '../modules/web'
import {
	assertObject,
	assertCollection,
	testWrapper,
} from '../lib/utility'

const PROPS = [
	'FileName',
	'ServerRelativeUrl'
]

const assertObjectProps = assertObject(PROPS)
const assertCollectionProps = assertCollection(PROPS)

const item = web('test/spx').list('Attachment').item(1)

export default {
	get: () => testWrapper('attachment GET')([
		() => assertObjectProps('attachment test.xlsx')(item.attachment('test.xlsx').get()),
		() => assertCollectionProps('attachment collection test.xlsx')(item.attachment(['test.xlsx']).get()),
	]),
	all() { testWrapper('attachment ALL')([this.get]) }
}
