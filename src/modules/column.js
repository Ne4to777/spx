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
	isArrayFilled,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	listReport,
	isStrictUrl,
	isFilled,
	deep1Iterator
} from '../lib/utility'

const addFieldAsXml = spParentObject => schema => spParentObject.addFieldAsXml(
	schema, true, SP.AddFieldOptions.defaultValue
)

const liftColumnType = switchCase(typeOf)({
	object: column => {
		const newColumn = Object.assign({}, column)
		if (!column.Url) newColumn.Url = column.Title
		if (column.Url !== '/') newColumn.Url = shiftSlash(newColumn.Url)
		if (!column.Title) newColumn.Title = column.EntityPropertyName || column.InternalName || column.StaticName
		if (!column.Type) newColumn.Type = 'Text'
		return newColumn
	},
	string: column => ({
		Title: column,
		Url: column === '/' ? '/' : shiftSlash(mergeSlashes(column)),
		Type: 'Text'
	}),
	default: () => ({
		Title: '',
		Url: '',
		Type: 'Text'
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value)
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([map(liftColumnType), removeEmptyUrls, removeDuplicatedUrls]),
				constant([liftColumnType()])
			])(value)
			: liftColumnType(value)
	}
}


class Column {
	constructor(parent, folders) {
		this.name = 'column'
		this.parent = parent
		this.box = getInstance(Box)(folders)
		this.contextUrl = parent.contextUrl
		this.getContextSPObject = parent.getContextSPObject
		this.getListSPObject = parent.getSPObject
		this.iterator = deep1Iterator({
			contextUrl: this.contextUrl,
			elementBox: this.box,
		})
	}

	async	get(opts) {
		const { listUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			console.log(listSPObject)
			const elementUrl = element.Url
			const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl)
			const spObject = isCollection
				? this.getSPObjectCollection(listSPObject)
				: this.getSPObject(elementUrl, listSPObject)
			return load(clientContext)(spObject)(opts)
		})
		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		return prepareResponseJSOM(opts)(result)
	}

	async	create(opts) {
		const { listUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const url = element.Title || element.Url
			if (!isStrictUrl(url)) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const {
				Title = url,
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
				addFieldAsXml(this.getSPObjectCollection(listSPObject)),
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
						set_title: element.Title,
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

			return load(clientContext)(spObject)(opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('create', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	update(opts) {
		const { listUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			if (!isStrictUrl(element.Url)) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
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
					set_title: element.Title,
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
			])(this.getSPObject(element.Url, listSPObject))
			return load(clientContext)(spObject)(opts)
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('update', opts)
		return prepareResponseJSOM(opts)(result)
	}

	async	delete(opts) {
		const { listUrl } = this
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const elementUrl = element.Url
			if (!isStrictUrl(elementUrl)) return undefined
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(listUrl, contextSPObject)
			const spObject = this.getSPObject(elementUrl, listSPObject)
			spObject.deleteObject()
			return elementUrl
		})
		if (this.box.getCount()) {
			await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))
		}
		this.report('delete', opts)
		return prepareResponseJSOM(opts)(result)
	}

	getSPObject(elementUrl, parentSPObject) {
		console.log(parentSPObject)
		const fields = parentSPObject.get_fields()
		return isGUID(elementUrl)
			? fields.getById(elementUrl)
			: fields.getByInternalNameOrTitle(elementUrl)
	}

	getSPObjectCollection(parentSPObject) {
		return parentSPObject.get_fields()
	}

	report(actionType, opts = {}) {
		listReport(actionType, {
			...opts,
			name: this.name,
			box: this.box,
			listBox: this.parent.box,
			contextBox: this.parent.parent.box
		})
	}

	of(columns) {
		return getInstance(this.constructor)(this.parent, columns)
	}
}
export default getInstance(Column)
