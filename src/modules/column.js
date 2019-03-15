import {
	ACTION_TYPES_TO_UNSET,
	ACTION_TYPES,
	Box,
	getInstance,
	isGUID,
	pipe,
	methodEmpty,
	method,
	ifThen,
	constant,
	prepareResponseJSOM,
	getClientContext,
	urlSplit,
	load,
	executorJSOM,
	setFields,
	overstep,
	stringReplace,
	isFilled,
	isExists,
	switchCase,
	prop,
	isStringEmpty,
	slice,
	hasUrlTailSlash
} from './../utility';
import * as cache from './../cache';

// Internal

const NAME = 'column';

const getSPObject = element => pipe([
	methodEmpty('get_fields'),
	ifThen(constant(isGUID(element)))([
		method('getById')(element),
		method('getByInternalNameOrTitle')(element)
	]),
])

const getSPObjectCollection = methodEmpty('get_fields');

const addFieldAsXml = spParentObject => schema => spParentObject.addFieldAsXml(schema, true, SP.AddFieldOptions.defaultValue);

const report = ({ silent, actionType }) => contextBox => parentBox => box => spObjects => (
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} in ${parentBox.join()} at ${contextBox.join()}`),
	spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => (opts = {}) => parent.parent.box
	.chainAsync(async contextElement => {
		const { cached } = opts;
		let needToQuery = true;
		const spObjectsToCache = new Map;
		const contextUrl = contextElement.Url;
		const clientContext = getClientContext(contextUrl);
		const contextUrls = urlSplit(contextUrl);
		const contextSPObject = parent.parent.getSPObject(clientContext);
		const spObjects = parent.box.chain(listElement =>
			box.chain(element => {
				const columnTitle = element.Title;
				const listUrl = listElement.Title;
				const listSPObject = parent.getSPObject(listUrl)(contextSPObject);
				const isCollection = !columnTitle || hasUrlTailSlash(columnTitle);
				const spObject = spObjectGetter({
					spParentObject: actionType === 'create' || isCollection ? getSPObjectCollection(listSPObject) : getSPObject(columnTitle)(listSPObject),
					element
				});
				const cachePaths = [...contextUrls, 'lists', listUrl, NAME, isCollection ? cacheLeaf + 'Collection' : cacheLeaf, columnTitle];
				ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -3)(cachePaths));
				const spObjectCached = cached ? cache.get(cachePaths) : null;
				if (actionType === 'delete' || actionType === 'recycle') return;
				if (cached && spObjectCached) {
					needToQuery = false;
					return spObjectCached;
				} else {
					const currentSPObjects = load(clientContext)(spObject)(opts);
					spObjectsToCache.set(cachePaths, currentSPObjects)
					return currentSPObjects;
				}
			}))
		if (needToQuery) {
			await executorJSOM(clientContext)(opts)
			spObjectsToCache.forEach((value, key) => cache.set(value)(key))
		};
		return spObjects;
	})
	.then(report({ ...opts, actionType })(parent.parent.box)(parent.box)(box))
	.then(prepareResponseJSOM(opts))


// Inteface

export default (parent, elements) => {
	const instance = {
		box: getInstance(Box)(elements, 'column'),
		parent
	};
	const executeBinded = execute(parent)(instance.box)('properties');
	return {

		get: executeBinded(null)(prop('spParentObject')),

		create: executeBinded('create')(({ spParentObject, element }) => {
			const {
				Title,
				Type = element.TypeAsString || 'Text',
				AllowMultipleValues,
				LookupWebId,
				LookupList,
				LookupField = 'Title',
				MaxLength,
				RichText,
				SchemaXml
			} = element;
			const clientContext = spParentObject.get_context();
			const castTo = value => spFieldObject => clientContext.castTo(spFieldObject, value);
			return pipe([
				ifThen(isFilled)([
					ifThen(constant(MaxLength))([
						stringReplace(/MaxLength="\d+"/)(`MaxLength="${MaxLength}"`)
					]),
					ifThen(constant(MaxLength && Type === 'Text'))([
						constant(`<Field Type="${Type}" DisplayName="${Title}" MaxLength="${MaxLength}"/>`),
						constant(`<Field Type="${Type}" DisplayName="${Title}"/>`)
					])
				]),
				addFieldAsXml(spParentObject),
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
		}),

		update: executeBinded('update')(({ spParentObject, element }) => pipe([
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
		])(spParentObject)),

		delete: opts => executeBinded('delete')(({ spParentObject }) =>
			overstep(methodEmpty('deleteObject'))(spParentObject))(opts)
	}
}