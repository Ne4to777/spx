/* eslint no-unused-vars:0 */
import { set, unset, inspect } from '../lib/cache'
import { assert, testWrapper } from '../lib/utility'

const testUnit = path => operation => sample => {
	if (operation) operation(path)
	const cacheStr = JSON.stringify(inspect())
	assert(`cache\nhave:${cacheStr}\nshould:${sample}`)(cacheStr === sample)
}

const testWithPath = testUnit(['a', 'b'])
export default {
	run: () => testWrapper('cache')([
		() => testWithPath()('{}'),
		() => testWithPath(set(1))('{"a":{"b":1}}'),
		() => testWithPath(unset)('{"a":{}}'),
		() => testWithPath(set(2))('{"a":{"b":2}}'),
		() => unset(['a'])
	])
}
