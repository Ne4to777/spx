import { pipe, filter, isArrayFilled, fix, ifThen, constant, isNotFilled, identity } from './../utility'

const cache = {};

const cacheCombinator = f => fix(f)(cache);
const filterNotEmpty = filter(isArrayFilled);

export const get = pipe([
  filterNotEmpty,
  cacheCombinator(fR => base => ([h, ...t]) => t.length ? (base[h] ? fR(base[h])(t) : null) : base[h])
]);

export const set = spObjects => pipe([
  filterNotEmpty,
  cacheCombinator(fR => base => ([h, ...t]) => t.length ? (!base[h] && (base[h] = {}), fR(base[h])(t)) : base[h] = spObjects),
  constant(spObjects),
])

export const unset = pipe([
  filterNotEmpty,
  cacheCombinator(fR => base => ([h, ...t]) => t.length ? (base[h] ? fR(base[h])(t) : void 0) : (base ? delete base[h] : void 0))
]);

export const inspect = _ => console.log(cache);

const test = _ => {
  const testWithPath = testUnit(['a', 'b']);
  testWithPath()('{}');
  testWithPath(set(1))('{"a":{"b":1}}');
  testWithPath(unset)('{"a":{}}');
  testWithPath(set(2))('{"a":{"b":2}}');
}

const testUnit = path => operation => sample => {
  operation && operation(path);
  const cacheStr = JSON.stringify(cache);
  cacheStr !== sample && log(`have:${cacheStr}\nshould:${sample}`)
}

// test()