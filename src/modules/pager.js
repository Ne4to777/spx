/* eslint prefer-const:0 */
import {
	getInstance,
	isArray,
	stringCut,
	arrayHead,
	arrayLast,
	isDefined
} from '../lib/utility'


const getPageValue = data => data === null
	? null
	: data instanceof Date
		? data.toISOString()
		: data.get_lookupValue
			? data.get_lookupValue()
			: data.toString
				? data.toString()
				: data

class Pager {
	constructor(parent, params = {}) {
		this.name = 'pager'
		this.parent = parent
		const { OrderBy } = params
		this.mainParams = {
			OrderBy,
			Limit: isDefined(params.Limit) ? params.Limit : 10,
			Scope: params.Scope,
			Query: params.Query,
			Folder: params.Folder,
		}

		this.onLoad = params.OnLoad
		this.onLoadForward = params.OnLoadForward
		this.onLoadBackward = params.OnLoadBackward
		this.onLoadAll = params.OnLoadAll
		this.onLoadAllForward = params.OnLoadAllForward
		this.onLoadAllBackward = params.OnLoadAllBackward
		this.onPreload = params.OnPreload
		this.onPreloadForward = params.OnPreloadForward
		this.onPreloadBackward = params.OnPreloadBackward

		this.itemIds = {
			first: 0,
			last: 0
		}
		this.itemColumns = {
			first: {},
			lats: {}
		}

		this.orderColumnTitles = []

		if (OrderBy) {
			const cut = stringCut(/>$/)
			this.orderColumnTitles = isArray(OrderBy) ? OrderBy.map(cut) : [cut(OrderBy)]
		}
	}

	async move(opts = {}) {
		let { view, groupBy, isPrevious } = opts
		if (isPrevious) {
			if (this.onPreloadBackward) await this.onPreloadBackward()
		} else if (this.onPreloadForward) await this.onPreloadForward()
		if (this.onPreload) await this.onPreload()
		if (view) view = [].concat(view, ...this.orderColumnTitles, 'ID')
		if (groupBy) groupBy = undefined

		let res = await this.parent
			.item({
				...this.mainParams,
				Page: {
					IsPrevious: isPrevious,
					ID: this.itemIds[isPrevious ? 'first' : 'last'],
					Columns: this.itemColumns[isPrevious ? 'first' : 'last']
				}
			})
			.get({ ...opts, view, groupBy })

		if (isPrevious) {
			if (this.onLoadBackward) await this.onLoadBackward()
		} else if (this.onLoadForward) await this.onLoadForward()
		if (this.onLoad) await this.onLoad()

		if (res.length) {
			const firstItem = arrayHead(res)
			const lastItem = arrayLast(res)
			if (isPrevious) {
				if (this.itemIds.first === firstItem.ID) {
					res = []
					this.itemIds.last = 0
					this.lastItemColumns = {}
					if (this.onLoadAll) await this.onLoadAll()
					if (this.onLoadAllBackward) await this.onLoadAllBackward()
				} else {
					this.itemIds.first = firstItem.ID
					this.itemIds.last = lastItem.ID
					if (this.orderColumnTitles.length) {
						this.itemColumns = this.orderColumnTitles.reduce((acc, el) => {
							acc.first[el] = getPageValue(firstItem[el])
							acc.last[el] = getPageValue(lastItem[el])
							return acc
						}, { first: {}, last: {} })
					}
				}
			} else {
				this.itemIds.first = firstItem.ID
				this.itemIds.last = lastItem.ID
				if (this.orderColumnTitles.length) {
					this.itemColumns = this.orderColumnTitles.reduce((acc, el) => {
						acc.first[el] = getPageValue(firstItem[el])
						acc.last[el] = getPageValue(lastItem[el])
						return acc
					}, { first: {}, last: {} })
				}
			}
		} else {
			if (isPrevious) {
				if (this.onLoadAllBackward) await this.onLoadAllBackward()
			} else if (this.onLoadAllForward) await this.onLoadAllForward()
			if (this.onLoadAll) await this.onLoadAll()
		}
		return res
	}

	next(opts) {
		return this.move(opts)
	}

	previous(opts = {}) {
		return this.move({ ...opts, isPrevious: true })
	}
}

export default getInstance(Pager)
