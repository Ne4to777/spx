import { set, unset, inspect } from './../cache'
import { assert, testIsOk } from './../utility'

const testUnit = path => operation => sample => {
  operation && operation(path);
  const cacheStr = JSON.stringify(inspect());
  assert(`cache\nhave:${cacheStr}\nshould:${sample}`)(cacheStr === sample)
}


export default _ => {
  const testWithPath = testUnit(['a', 'b']);
  testWithPath()('{}');
  testWithPath(set(1))('{"a":{"b":1}}');
  testWithPath(unset)('{"a":{}}');
  testWithPath(set(2))('{"a":{"b":2}}');
  testIsOk('cache')()
}