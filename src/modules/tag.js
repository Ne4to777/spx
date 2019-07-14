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
	deep2Iterator,
	ifThen,
	isArrayFilled,
	pipe,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	constant,
	setFields,
	getInstanceEmpty,
	webReport
} from '../lib/utility'

// Internal

const NAME = 'tag'

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

const iterator = instance => deep2Iterator({
	contextBox: instance.parent.box,
	elementBox: instance.box
})

const get = instance => isExact => async opts => {
	const { clientContexts, result } = await iterator(instance)(({ clientContext, element }) => {
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
	await instance.parent.box.chain(el => Promise.all(
		clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))
	))
	return prepareResponseJSOM(opts)(isExact && result.length === 1 ? result[0] : result)
}

// Interface

export default parent => urls => {
	const instance = {
		box: getInstance(Box)(urls),
		NAME,
		parent
	}
	const iteratorInstanced = iterator(instance)
	const report = actionType => (opts = {}) => webReport({
		...opts, NAME, actionType, box: instance.box, contextBox: instance.parent.box
	})
	return {
		get: get(instance)(true),
		search: get(instance)(),
		create: async opts => {
			const { clientContexts, result } = await iteratorInstanced(({ clientContext, element }) => {
				const elementUrl = element.Url
				if (!elementUrl) return undefined
				const spObject = getTermSet(clientContext).createTerm(
					elementUrl,
					1033,
					getInstanceEmpty(SP.Guid.newGuid).toString()
				)
				return load(clientContext)(spObject)(opts)
			})
			if (instance.box.getCount()) {
				await instance.parent.box.chain(el => Promise.all(
					clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))
				))
			}
			report('create')(opts)
			return prepareResponseJSOM(opts)(result)
		},

		update: async opts => {
			const { clientContexts, result } = await iteratorInstanced(({ clientContext, element }) => {
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
			if (instance.box.getCount()) {
				await instance.parent.box.chain(
					el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
				)
			}
			report('update')(opts)
			return prepareResponseJSOM(opts)(result)
		},

		delete: async opts => {
			const { clientContexts, result } = await iteratorInstanced(({ clientContext, element }) => {
				const elementUrl = element.Url
				if (!elementUrl) return undefined
				const termStore = getTermStore(clientContext)
				const spObject = getSPObject(clientContext)(elementUrl)
				methodEmpty('deleteObject')(spObject)
				methodEmpty('commitAll')(termStore)
				return elementUrl
			})
			if (instance.box.getCount()) {
				await instance.parent.box.chain(
					el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts)))
				)
			}
			report('delete')(opts)
			return prepareResponseJSOM(opts)(result)
		}
	}
}
