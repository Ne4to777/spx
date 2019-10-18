
import web from '../modules/web'
import {
	assertObject,
	testWrapper,
	assert
} from '../lib/utility'

const PROPS = [
	'Description',
	'Id',
	'Information'
]

const assertProps = assertObject(PROPS)

export default {
	get: () => testWrapper('time GET')([
		() => web()
			.time()
			.get(time => assert('time is wrong')(time.format('dd.MM.yyyy') === new Date().format('dd.MM.yyyy'))),
		() => assertProps('time zone')(web().time().getZone()),
	]),
	all: () => testWrapper('time ALL')([this.get])
}
