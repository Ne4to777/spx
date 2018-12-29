import * as utility from './../utility';
import * as cache from './../cache';

export default class Column {
	constructor(parent, elementUrl) {
		this._elementUrl = elementUrl;
		this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
		this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl || ''];
		this._parent = parent;
		this._contextUrlIsArray = this._parent._parent._contextUrlIsArray;
		this._contextUrls = this._parent._parent._contextUrls;
		this._listUrlIsArray = this._parent._elementUrlIsArray;
		this._listUrls = this._parent._elementUrls;
	}

	// Inteface

	async get(opts) {
		return this._execute(null, spObject => (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
	}

	async create(opts) {
		return this._execute('create', (spContextObject, element) => {
			let spObject;
			const {
				Title,
				TypeAsString,
				Type = 'Text',
				AllowMultipleValues,
				LookupWebId,
				LookupList,
				LookupField = 'Title',
				MaxLength,
				RichText,
				SchemaXml
			} = element;
			const type = TypeAsString || Type;
			const clientContext = spContextObject.get_context();
			let schema = SchemaXml;
			if (SchemaXml) {
				if (MaxLength) schema = SchemaXml.replace(/MaxLegth="\d+"/);
			} else {
				schema = MaxLength && type === 'Text' ?
					`<Field Type="${type}" DisplayName="${Title}" MaxLength="${MaxLength}"/>` :
					`<Field Type="${type}" DisplayName="${Title}"/>`;
			}
			const spFieldObject = spContextObject.addFieldAsXml(schema, true, SP.AddFieldOptions.defaultValue);
			utility.setFields(spFieldObject, {
				set_defaultValue: element.DefaultValue,
				set_description: element.Description,
				set_direction: element.Direction,
				set_enforceUniqueValues: element.EnforceUniqueValues,
				set_fieldTypeKind: element.FieldTypeKind,
				set_group: element.Group,
				// set_hidden: element.Hidden ,
				set_indexed: element.Indexed,
				set_jsLink: element.JsLink,
				set_objectVersion: element.ObjectVersion,
				set_readOnlyField: element.ReadOnlyField,
				set_required: element.Required,
				set_schemaXml: element.SchemaXml ? element.SchemaXml.replace(/ID="{[^}]+}"/, '') : '',
				set_staticName: element.StaticName,
				set_title: element.Title,
				set_typeAsString: element.TypeAsString,
				set_validationFormula: element.ValidationFormula || void 0,
				set_validationMessage: element.ValidationMessage || void 0
			});
			const castTo = value => clientContext.castTo(spFieldObject, value);

			switch (type) {
				case 'Text': spObject = castTo(SP.FieldText); break;
				case 'Note':
					spObject = castTo(SP.FieldMultiLineText);
					RichText && spObject.set_richText(true);
					break;
				case 'Likes':
				case 'Number': spObject = castTo(SP.FieldNumber); break;
				case 'Boolean': spObject = castTo(SP.Field); break;
				case 'Choice': spObject = castTo(AllowMultipleValues ? SP.FieldMultiChoice : SP.FieldChoice); break;
				case 'DateTime': spObject = castTo(SP.FieldDateTime); break;
				case 'URL': spObject = castTo(SP.FieldUrl); break;
				case 'RatingCount':
				case 'AverageRating': spObject = castTo(SP.FieldRatingScale); break;
				case 'Lookup':
				case 'LookupMulti':
					spObject = castTo(SP.FieldLookup);
					utility.setField(spObject, 'set_lookupWebId', LookupWebId);
					spObject.set_lookupList(LookupList);
					spObject.set_lookupField(LookupField);
					AllowMultipleValues && spObject.set_allowMultipleValues(true);
					break;
				case 'User':
				case 'UserMulti':
					spObject = castTo(SP.FieldUser);
					AllowMultipleValues && spObject.set_allowMultipleValues(true);
					break;
			}
			if (spObject) {
				spObject.update();
				spObject.cachePath = 'property';
				return spObject;
			}
		}, opts);
	}

	async update(opts) {
		return this._execute('update', (spObject, element) => {
			utility.setFields(spObject, {
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
			})
			if (element.MaxLength) {
				const name = element.Name || element.Title;
				spObject.set_schemaXml(`<Field Type="Text" DisplayName="${name}" MaxLength="${element.MaxLength}" />`);
			}
			spObject.update();
			spObject.cachePath = 'property';
			return spObject;
		}, opts);
	}

	async delete(opts) {
		return this._execute('delete', spObject => (spObject.deleteObject(), spObject.cachePath = 'property', spObject), opts)
	}

	// Internal

	get _name() { return 'column' }

	async _execute(actionType, spObjectGetter, opts = {}) {
		const { cached } = opts;
		const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
			let needToQuery;
			const spObjectsToCache = new Map;
			const clientContext = utility.getClientContext(contextUrl);
			const contextUrls = contextUrl.split('/');
			const spObjects = await Promise.all(this._listUrls.reduce((acc, listUrl) => acc.concat(this._elementUrls.map(elementUrl => {
				const element = this._liftElementUrlType(elementUrl);
				let fieldTitle = element.Name || element.Title;
				if (actionType === 'create') fieldTitle = '';
				const spObject = spObjectGetter(this._getSPObject(clientContext, listUrl, fieldTitle), element);
				const cachePaths = [...contextUrls, listUrl, fieldTitle, this._name, spObject.cachePath];
				utility.ACTION_TYPES_TO_UNSET[actionType] && cache.unset([...contextUrls, listUrl]);
				if (actionType === 'delete' || actionType === 'recycle') {
					needToQuery = true;
				} else {
					const spObjectCached = cached ? cache.get(cachePaths) : null;
					if (cached && spObjectCached) {
						return spObjectCached;
					} else {
						needToQuery = true;
						const currentSPObjects = utility.load(clientContext, spObject, opts);
						spObjectsToCache.set(cachePaths, currentSPObjects)
						return currentSPObjects;
					}
				}
			})), []))

			if (needToQuery) {
				await utility.executeQueryAsync(clientContext, opts);
				spObjectsToCache.forEach((value, key) => cache.set(key, value))
			};
			return spObjects;
		}))

		this._log(actionType, opts);

		return utility.prepareResponseJSOM(elements, opts);
	}

	_getSPObject(clientContext, listUrl, elementUrl) {
		const columns = this._parent._getSPObject(clientContext, listUrl).get_fields();
		return elementUrl ? (columns[utility.isGUID(elementUrl) ? 'getById' : 'getByInternalNameOrTitle'](elementUrl)) : columns
	}

	_liftElementUrlType(elementUrl) {
		switch (typeOf(elementUrl)) {
			case 'object':
				return elementUrl;
			case 'string':
				return {
					Title: elementUrl,
					Name: elementUrl
				}
		}
	}

	_log(actionType, opts = {}) {
		!opts.silent && actionType &&
			console.log(`${
				utility.ACTION_TYPES[actionType]} ${
				this._name} at ${
				this._contextUrls.join(', ')}: ${
				this._elementUrls.map(el => this._liftElementUrlType(el).Title).join(', ')} in ${
				this._listUrls.join(', ')}`);
	}
}