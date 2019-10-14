/* eslint-disable class-methods-use-this */
import {
	AbstractBox,
	load,
	methodEmpty,
	getInstance,
	prepareResponseJSOM,
	executorJSOM,
	switchType,
	shiftSlash,
	pipe,
	removeEmptiesByProp,
	removeDuplicatedProp,
	setFields,
	getInstanceEmpty,
	rootReport,
	deep1Iterator
} from '../lib/utility'

const KEY_PROP = 'Label'

const getTermStore = clientContext => SP
	.Taxonomy
	.TaxonomySession
	.getTaxonomySession(clientContext)
	.getDefaultKeywordsTermStore()

const getTermSet = clientContext => getTermStore(clientContext).get_keywordsTermSet()


const getAllTerms = clientContext => getTermSet(clientContext).getAllTerms()


const arrayValidator = pipe([removeEmptiesByProp(KEY_PROP), removeDuplicatedProp(KEY_PROP)])

const lifter = switchType({
	object: tag => {
		const newTag = Object.assign({}, tag)
		if (tag[KEY_PROP] !== '/') newTag[KEY_PROP] = shiftSlash(newTag[KEY_PROP])
		return newTag
	},
	string: tag => {
		const tagName = tag === '/' || !tag ? undefined : tag
		return {
			[KEY_PROP]: tagName
		}
	},
	default: () => ({
		[KEY_PROP]: undefined
	})
})

class Box extends AbstractBox {
	constructor(value) {
		super(value, lifter, arrayValidator)
		this.joinProp = KEY_PROP
		this.prop = KEY_PROP
	}
}

class Tag {
	constructor(parent, tags) {
		this.name = 'tag'
		this.parent = parent
		this.tags = tags
		this.box = getInstance(Box)(tags)
		this.count = this.box.getCount()
		this.iterator = deep1Iterator({ elementBox: this.box })
	}

	async get(opts = {}) {
		const { isExact = true } = opts
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const spObject = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext)
				.getDefaultKeywordsTermStore()
				.get_keywordsTermSet()
				.getTerms(setFields({
					set_defaultLabelOnly: element.DefaultLabelOnly || false,
					// set_excludeKeyword: element.ExcludeKeyword || false,
					set_resultCollectionSize: element.ResultCollectionSize || element.Limit || 100000,
					set_stringMatchOption: SP.Taxonomy.StringMatchOption[isExact ? 'exactMatch' : 'startsWith'],
					set_termLabel: element[KEY_PROP],
					set_trimDepricated: element.TrimDepricated || true,
					set_trimUnavailable: element.TrimUnavailable || true
				})(SP.Taxonomy.LabelMatchInformation.newObject(clientContext)))
			return load(clientContext, spObject, opts)
		})

		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(isExact && this.box.isArray === 1 ? result[0] : result, opts)
	}

	async search(opts) {
		return this.get({ ...opts, isExact: false })
	}

	async create(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementLabel = element[KEY_PROP]
			if (!elementLabel) return undefined
			const spObject = getTermSet(clientContext).createTerm(
				elementLabel,
				1033,
				getInstanceEmpty(SP.Guid.newGuid).toString()
			)
			return load(clientContext, spObject, opts)
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	async update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementLabel = element[KEY_PROP]
			if (!elementLabel) return undefined
			const spObject = this.getSPObject(elementLabel, clientContext)
			setFields({
				set_isAvailableForTagging: element.IsAvailableForTagging,
				set_name: element.Name,
				set_owner: element.Owner
			})(spObject)
			return load(clientContext, spObject, opts)
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
	}

	async delete(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementLabel = element[KEY_PROP]
			if (!elementLabel) return undefined
			const termStore = getTermStore(clientContext)
			const spObject = this.getSPObject(elementLabel, clientContext)
			methodEmpty('deleteObject')(spObject)
			methodEmpty('commitAll')(termStore)
			return elementLabel
		})
		if (this.count) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('delete', opts)
		return prepareResponseJSOM(result, opts)
	}

	report(actionType, opts = {}) {
		rootReport(actionType,
			{
				...opts,
				name: this.name,
				box: this.box
			})
	}

	getSPObject(elementUrl, clientContext) {
		return getAllTerms(clientContext).getByName(elementUrl)
	}
}

export default getInstance(Tag)
