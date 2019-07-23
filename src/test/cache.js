/* eslint no-unused-vars:0 */
import { set, unset, inspect } from '../lib/cache'
import { assert, testIsOk } from '../lib/utility'

const testUnit = path => operation => sample => {
	if (operation) operation(path)
	const cacheStr = JSON.stringify(inspect())
	assert(`cache\nhave:${cacheStr}\nshould:${sample}`)(cacheStr === sample)
}

export default () => {
	const testWithPath = testUnit(['a', 'b'])
	testWithPath()('{}')
	testWithPath(set(1))('{"a":{"b":1}}')
	testWithPath(unset)('{"a":{}}')
	testWithPath(set(2))('{"a":{"b":2}}')
	testIsOk('cache')()
}
