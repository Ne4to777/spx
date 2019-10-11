/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	getInstance,
	isGUID,
	pipe,
	methodEmpty,
	method,
	ifThen,
	constant,
	prepareResponseJSOM,
	load,
	executorJSOM,
	setFields,
	overstep,
	stringReplace,
	isExists,
	switchCase,
	isStringEmpty,
	hasUrlTailSlash,
	typeOf,
	shiftSlash,
	mergeSlashes,
	listReport,
	isStrictUrl,
	isFilled,
	deep1Iterator,
	removeEmptiesByProp,
	removeDuplicatedProp
} from '../lib/utility'

const KEY_PROP = 'Title'

const addFieldAsXml = spParentObject => schema => spParentObject.addFieldAsXml(
	schema, true, SP.AddFieldOptions.defaultValue
)

const arrayValidator = pipe([removeEmptiesByProp(KEY_PROP), removeDuplicatedProp(KEY_PROP)])

const lifter = switchCase(typeOf)({
	object: column => {
		const newColumn = Object.assign({}, column)
		if (column[KEY_PROP] !== '/') newColumn[KEY_PROP] = shiftSlash(newColumn[KEY_PROP])
		if (!column[KEY_PROP]) {
			newColumn[KEY_PROP] = column.EntityPropertyName
				|| column.InternalName
				|| column.StaticName
		}
		if (!column.Type) newColumn.Type = 'Text'
		return newColumn
	},
	string: column => ({
		[KEY_PROP]: column === '/' ? '/' : shiftSlash(mergeSlashes(column)),
		Type: 'Text'
	}),
	default: () => ({
		[KEY_PROP]: undefined,
		Type: 'Text'
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value, lifter, arrayValidator)
		this.prop = KEY_PROP
		this.joinProp = KEY_PROP
	}
}


class Column {
	constructor(parent, columns) {
		this.name = 'column'
		this.parent = parent
		this.box = getInstance(Box)(columns)
		this.contextUrl = parent.parent.box.getHeadPropValue()
		this.listUrl = parent.box.getHeadPropValue()
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box,
		})
	}

	async	get(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementTitle = element[KEY_PROP]
			const isCollection = isStringEmpty(elementTitle) || hasUrlTailSlash(elementTitle)
			const spObject = isCollection
				? this.getSPObjectCollection(clientContext)
				: this.getSPObject(elementTitle, clientContext)
			return load(clientContext, spObject, opts)
		})
		await Promise.all(clientContexts.map(executorJSOM))
		return prepareResponseJSOM(result, opts)
	}

	async	create(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const title = element[KEY_PROP]
			if (!isStrictUrl(title)) return undefined
			const {
				Title = title,
				Type = element.TypeAsString || 'Text',
				AllowMultipleValues,
				LookupWebId,
				LookupList,
				LookupField = 'Title',
				MaxLength,
				RichText,
				SchemaXml
			} = element
			const castTo = value => spFieldObject => clientContext.castTo(spFieldObject, value)
			const spObject = pipe([
				ifThen(isFilled)([
					ifThen(constant(MaxLength))([stringReplace(/MaxLength="\d+"/)(`MaxLength="${MaxLength}"`)]),
					ifThen(constant(MaxLength && Type === 'Text'))([
						constant(`<Field Type="${Type}" DisplayName="${Title}" MaxLength="${MaxLength}"/>`),
						constant(`<Field Type="${Type}" DisplayName="${Title}"/>`)
					])
				]),
				addFieldAsXml(this.getSPObjectCollection(clientContext)),
				overstep(
					setFields({
						set_defaultValue: element.DefaultValue,
						set_description: element.Description,
						set_direction: element.Direction,
						set_enforceUniqueValues: element.EnforceUniqueValues,
						set_fieldTypeKind: element.FieldTypeKind,
						set_group: element.Group,
						set_hidden: element.Hidden || undefined,
						set_indexed: element.Indexed,
						set_jsLink: element.JsLink,
						set_objectVersion: element.ObjectVersion,
						set_readOnlyField: element.ReadOnlyField,
						set_required: element.Required,
						set_schemaXml: element.SchemaXml
							? element.SchemaXml.replace(/\sID="{[^}]+}"/, '')
							: undefined,
						set_staticName: element.StaticName,
						set_title: element[KEY_PROP],
						set_typeAsString: element.TypeAsString,
						set_validationFormula: element.ValidationFormula || undefined,
						set_validationMessage: element.ValidationMessage || undefined
					})
				),
				switchCase(constant(Type))({
					Text: castTo(SP.FieldText),
					Note: pipe([
						castTo(SP.FieldMultiLineText),
						overstep(ifThen(constant(RichText))([method('set_richText')(true)]))
					]),
					Likes: castTo(SP.FieldNumber),
					Number: castTo(SP.FieldNumber),
					Boolean: castTo(SP.Field),
					Choice: castTo(AllowMultipleValues ? SP.FieldMultiChoice : SP.FieldChoice),
					DateTime: castTo(SP.FieldDateTime),
					URL: castTo(SP.FieldUrl),
					RatingCount: castTo(SP.FieldRatingScale),
					AverageRating: castTo(SP.FieldRatingScale),
					Lookup: pipe([
						castTo(SP.FieldLookup),
						overstep(
							pipe([
								method('set_lookupWebId')(LookupWebId),
								method('set_lookupList')(LookupList),
								method('set_lookupField')(LookupField),
								ifThen(constant(AllowMultipleValues))([method('set_allowMultipleValues')(true)])
							])
						)
					]),
					LookupMulti: pipe([
						castTo(SP.FieldLookup),
						overstep(
							pipe([
								method('set_lookupWebId')(LookupWebId),
								method('set_lookupList')(LookupList),
								method('set_lookupField')(LookupField),
								method('set_allowMultipleValues')(true)
							])
						)
					]),
					User: pipe([
						castTo(SP.FieldUser),
						overstep(ifThen(constant(AllowMultipleValues))([method('set_allowMultipleValues')(true)]))
					]),
					UserMulti: pipe([castTo(SP.FieldUser), overstep(method('set_allowMultipleValues')(true))])
				}),
				overstep(ifThen(isExists)([methodEmpty('update')]))
			])(SchemaXml)

			return load(clientContext, spObject, opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('create', opts)
		return prepareResponseJSOM(result, opts)
	}

	async	update(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			if (!isStrictUrl(element[KEY_PROP])) return undefined
			const { MaxLength, Title } = element
			const spObject = pipe([
				setFields({
					set_defaultValue: element.DefaultValue,
					set_description: element.Description,
					set_direction: element.Direction,
					set_enforceUniqueValues: element.EnforceUniqueValues,
					set_fieldTypeKind: element.FieldTypeKind,
					set_group: element.Group,
					set_hidden: element.Hidden,
					set_indexed: element.Indexed,
					set_jsLink: element.JsLink,
					set_objectVersion: element.ObjectVersion,
					set_readOnlyField: element.ReadOnlyField,
					set_required: element.Required,
					set_schemaXml: element.SchemaXml,
					set_staticName: element.StaticName,
					set_title: element[KEY_PROP],
					set_typeAsString: element.TypeAsString,
					set_validationFormula: element.ValidationFormula,
					set_validationMessage: element.ValidationMessage
				}),
				overstep(
					pipe([
						ifThen(constant(element.MaxLength))([
							method('set_schemaXml')(
								`<Field Type="Text" DisplayName="${Title}" MaxLength="${MaxLength}"/>`
							)
						]),
						methodEmpty('update')
					])
				)
			])(this.getSPObject(element[KEY_PROP], clientContext))
			return load(clientContext, spObject, opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('update', opts)
		return prepareResponseJSOM(result, opts)
	}

	async	delete(opts) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementTitle = element[KEY_PROP]
			if (!isStrictUrl(elementTitle)) return undefined
			const spObject = this.getSPObject(elementTitle, clientContext)
			spObject.deleteObject()
			return elementTitle
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(executorJSOM))
		}
		this.report('delete', opts)
		return prepareResponseJSOM(result, opts)
	}

	getSPObject(elementTitle, clientContext) {
		return this.getSPObjectCollection(clientContext)[
			isGUID(elementTitle)
				? 'getById'
				: 'getByInternalNameOrTitle'
		](elementTitle)
	}

	getSPObjectCollection(clientContext) {
		return this.parent.getSPObject(this.listUrl, clientContext).get_fields()
	}

	report(actionType, opts = {}) {
		listReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			listUrl: this.listUrl,
			contextUrl: this.contextUrl
		})
	}

	of(columns) {
		return getInstance(this.constructor)(this.parent, columns)
	}
}
export default getInstance(Column)
