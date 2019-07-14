import {
	pipe, filter, fix, constant, isExists
} from './utility'

const cache = {}

const cacheCombinator = f => fix(f)(cache)
const filterNotEmpty = filter(isExists)

export const get = pipe([
	filterNotEmpty,
	cacheCombinator(fR => base => ([h, ...t]) => (t.length ? (base[h] ? fR(base[h])(t) : null) : base[h]))
])

/* eslint no-param-reassign:1 */
export const set = spObjects => pipe([
	filterNotEmpty,
	cacheCombinator(fR => base => ([h, ...t]) => {
		if (t.length) {
			if (!base[h]) base[h] = {}
			return fR(base[h])(t)
		}
		base[h] = spObjects
		return spObjects
	}),
	constant(spObjects)
])

export const unset = pipe([
	filterNotEmpty,
	cacheCombinator(
		fR => base => ([h, ...t]) => {
			if (t.length) {
				return base[h] ? fR(base[h])(t) : undefined
			}
			return base ? delete base[h] : undefined
		}
	)
])

export const inspect = () => cache
