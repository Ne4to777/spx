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
	deep3Iterator,
	listReport,
	isStrictUrl,
	isFilled
} from './../lib/utility';

// Internal

const NAME = 'column';

const getSPObject = elementUrl => pipe([
	methodEmpty('get_fields'),
	ifThen(constant(isGUID(elementUrl)))([
		method('getById')(elementUrl),
		method('getByInternalNameOrTitle')(elementUrl)
	]),
])

const getSPObjectCollection = methodEmpty('get_fields');

const liftColumnType = switchCase(typeOf)({
	object: column => {
		const newColumn = Object.assign({}, column);
		if (!column.Url) newColumn.Url = column.Title;
		if (column.Url !== '/') newColumn.Url = shiftSlash(newColumn.Url);
		if (!column.Title) newColumn.Title = column.EntityPropertyName || column.InternalName || column.StaticName;
		if (!column.Type) newColumn.Type = 'Text';
		return newColumn
	},
	string: column => ({
		Title: column,
		Url: column === '/' ? '/' : shiftSlash(mergeSlashes(column)),
		Type: 'Text'
	}),
	default: _ => ({
		Title: '',
		Url: '',
		Type: 'Text'
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value);
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([
					map(liftColumnType),
					removeEmptyUrls,
					removeDuplicatedUrls
				]),
				constant([liftColumnType()])
			])(value)
			: liftColumnType(value);
	}
}

const addFieldAsXml = spParentObject => schema => spParentObject.addFieldAsXml(schema, true, SP.AddFieldOptions.defaultValue);


// Inteface

export default parent => elements => {
	const instance = {
		box: getInstance(Box)(elements),
		parent
	};
	const iterator = deep3Iterator({
		contextBox: instance.parent.parent.box,
		parentBox: instance.parent.box,
		elementBox: instance.box
	});
	const report = actionType => opts => listReport({ ...opts, NAME, box: instance.box, listBox: instance.parent.box, contextBox: instance.parent.parent.box, actionType });
	return {

		get: async opts => {
			const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const elementUrl = element.Url;
				const isCollection = isStringEmpty(elementUrl) || hasUrlTailSlash(elementUrl);
				const spObject = isCollection
					? getSPObjectCollection(listSPObject)
					: getSPObject(elementUrl)(listSPObject);
				return load(clientContext)(spObject)(opts);
			})
			await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			return prepareResponseJSOM(opts)(result)
		},

		create: async opts => {
			const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
				const url = element.Title || element.Url;
				if (!isStrictUrl(url)) return;
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
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
				} = element;
				const castTo = value => spFieldObject => clientContext.castTo(spFieldObject, value);
				const spObject = pipe([
					ifThen(isFilled)([
						ifThen(constant(MaxLength))([
							stringReplace(/MaxLength="\d+"/)(`MaxLength="${MaxLength}"`)
						]),
						ifThen(constant(MaxLength && Type === 'Text'))([
							constant(`<Field Type="${Type}" DisplayName="${Title}" MaxLength="${MaxLength}"/>`),
							constant(`<Field Type="${Type}" DisplayName="${Title}"/>`)
						])
					]),
					addFieldAsXml(getSPObjectCollection(listSPObject)),
					overstep(setFields({
						set_defaultValue: element.DefaultValue,
						set_description: element.Description,
						set_direction: element.Direction,
						set_enforceUniqueValues: element.EnforceUniqueValues,
						set_fieldTypeKind: element.FieldTypeKind,
						set_group: element.Group,
						set_hidden: element.Hidden || void 0,
						set_indexed: element.Indexed,
						set_jsLink: element.JsLink,
						set_objectVersion: element.ObjectVersion,
						set_readOnlyField: element.ReadOnlyField,
						set_required: element.Required,
						set_schemaXml: element.SchemaXml ? element.SchemaXml.replace(/\sID\="{[^}]+}"/, '') : void 0,
						set_staticName: element.StaticName,
						set_title: element.Title,
						set_typeAsString: element.TypeAsString,
						set_validationFormula: element.ValidationFormula || void 0,
						set_validationMessage: element.ValidationMessage || void 0
					})),
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
							overstep(pipe([
								method('set_lookupWebId')(LookupWebId),
								method('set_lookupList')(LookupList),
								method('set_lookupField')(LookupField),
								ifThen(constant(AllowMultipleValues))([method('set_allowMultipleValues')(true)])
							]))
						]),
						LookupMulti: pipe([
							castTo(SP.FieldLookup),
							overstep(pipe([
								method('set_lookupWebId')(LookupWebId),
								method('set_lookupList')(LookupList),
								method('set_lookupField')(LookupField),
								method('set_allowMultipleValues')(true)
							]))
						]),
						User: pipe([
							castTo(SP.FieldUser),
							overstep(ifThen(constant(AllowMultipleValues))([method('set_allowMultipleValues')(true)]))
						]),
						UserMulti: pipe([
							castTo(SP.FieldUser),
							overstep(method('set_allowMultipleValues')(true))
						])
					}),
					overstep(ifThen(isExists)([methodEmpty('update')]))
				])(SchemaXml)

				return load(clientContext)(spObject)(opts);
			})
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			}
			report('create')(opts);
			return prepareResponseJSOM(opts)(result)
		},

		update: async opts => {
			const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
				if (!isStrictUrl(element.Url)) return;
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);

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
					overstep(pipe([
						ifThen(constant(element.MaxLength))([
							method('set_schemaXml')(`<Field Type="Text" DisplayName="${element.Title}" MaxLength="${element.MaxLength}" />`)
						]),
						methodEmpty('update')
					]))
				])(getSPObject(element.Url)(listSPObject))
				return load(clientContext)(spObject)(opts);
			})
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			}
			report('update')(opts);
			return prepareResponseJSOM(opts)(result)
		},

		delete: async opts => {
			const { clientContexts, result } = await iterator(({ clientContext, parentElement, element }) => {
				const elementUrl = element.Url;
				if (!isStrictUrl(elementUrl)) return;
				const contextSPObject = instance.parent.parent.getSPObject(clientContext);
				const listSPObject = instance.parent.getSPObject(parentElement.Url)(contextSPObject);
				const spObject = getSPObject(elementUrl)(listSPObject)
				spObject.deleteObject();
			})
			if (instance.box.getCount()) {
				await instance.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			}
			report('delete')(opts);
			return prepareResponseJSOM(opts)(result)
		},
	}
}