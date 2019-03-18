import { pipe, filter, fix, constant, isExists } from './utility'

const cache = {};

const cacheCombinator = f => fix(f)(cache);
const filterNotEmpty = filter(isExists);

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

export const inspect = _ => cache;