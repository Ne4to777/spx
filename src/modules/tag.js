import {
	AbstractBox,
	load,
	methodEmpty,
	getInstance,
	prepareResponseJSOM,
	executorJSOM,
	switchCase,
	typeOf,
	shiftSlash,
	ifThen,
	isArrayFilled,
	pipe,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	constant,
	setFields,
	getInstanceEmpty,
	webReport,
	deep1Iterator
} from '../lib/utility'

const getTermStore = clientContext => SP
	.Taxonomy
	.TaxonomySession
	.getTaxonomySession(clientContext)
	.getDefaultKeywordsTermStore()

const getTermSet = clientContext => getTermStore(clientContext).get_keywordsTermSet()


const getAllTerms = clientContext => getTermSet(clientContext).getAllTerms()

const getSPObject = clientContext => elementUrl => getAllTerms(clientContext).getByName(elementUrl)

const liftTagType = switchCase(typeOf)({
	object: tag => {
		const newTag = Object.assign({}, tag)
		if (!tag.Url) newTag.Url = tag.Name
		if (tag.Url !== '/') newTag.Url = shiftSlash(newTag.Url)
		if (!tag.Name) newTag.Name = tag.Url
		return newTag
	},
	string: tag => {
		const tagName = tag === '/' || !tag ? undefined : tag
		return {
			Url: tagName,
			Name: tagName
		}
	},
	default: () => ({
		Url: undefined,
		Name: undefined
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value)
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([map(liftTagType), removeEmptyUrls, removeDuplicatedUrls]),
				constant([liftTagType()])
			])(value)
			: liftTagType(value)
	}
}


const get = iterator => isExact => async opts => {
	const { clientContexts, result } = await iterator(({ clientContext, element }) => {
		const lmi = setFields({
			set_defaultLabelOnly: element.DefaultLabelOnly || false,
			// set_excludeKeyword: element.ExcludeKeyword || false,
			set_resultCollectionSize: element.ResultCollectionSize || element.Limit || 100000,
			set_stringMatchOption: SP.Taxonomy.StringMatchOption[isExact ? 'exactMatch' : 'startsWith'],
			set_termLabel: element.Url,
			set_trimDepricated: element.TrimDepricated || true,
			set_trimUnavailable: element.TrimUnavailable || true
		})(SP.Taxonomy.LabelMatchInformation.newObject(clientContext))

		const spObject = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext)
			.getDefaultKeywordsTermStore()
			.get_keywordsTermSet()
			.getTerms(lmi)
		return load(clientContext)(spObject)(opts)
	})

	await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
	return prepareResponseJSOM(opts)(isExact && result.length === 1 ? result[0] : result)
}

class Tag {
	constructor(tags) {
		this.tags = tags
		this.box = getInstance(Box)(tags)
		this.iterator = deep1Iterator({ elementBox: this.box })
	}

	async get(opts) {
		return get(this.iterator)(true)(opts)
	}

	async search(opts) {
		return get(this.iterator)(false)(opts)
	}

	async	create(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = element.Url
			if (!elementUrl) return undefined
			const spObject = getTermSet(clientContext).createTerm(
				elementUrl,
				1033,
				getInstanceEmpty(SP.Guid.newGuid).toString()
			)
			return load(clientContext)(spObject)(opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('create', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = element.Url
			if (!elementUrl) return undefined
			const spObject = getSPObject(clientContext)(elementUrl)
			setFields({
				set_isAvailableForTagging: element.IsAvailableForTagging,
				set_name: element.Name,
				set_owner: element.Owner
			})(spObject)
			return load(clientContext)(spObject)(opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('update', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async delete(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = element.Url
			if (!elementUrl) return undefined
			const termStore = getTermStore(clientContext)
			const spObject = getSPObject(clientContext)(elementUrl)
			methodEmpty('deleteObject')(spObject)
			methodEmpty('commitAll')(termStore)
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('delete', opts)
		return prepareResponseJSOM(opts)(result)
	}

	report(actionType, opts = {}) {
		webReport(actionType, { ...opts, NAME: 'tag', box: this.box })
	}
}

export default getInstance(Tag)
