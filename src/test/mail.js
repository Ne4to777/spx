/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertObject,
	testWrapper,
} from '../lib/utility'

const PROPS = [
	'SendEmail'
]

const assertProps = assertObject(PROPS)

export default {
	get: () => testWrapper('mail Send')([
		() => assertProps('mail send')(web('test/spx').mail({
			To: 'AlekseyAlekseew@dme.ru',
			Subject: 'test',
			Body: 'test'
		}).send()),
	]),
	all() {
		testWrapper('mail ALL')([
			this.get,
		])
	}
}
