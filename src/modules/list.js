import * as utility from './../utility';
import * as cache from './../cache';
import Column from './../modules/column';
import Folder from './../modules/folderList';
import File from './../modules/fileList';
import Item from './../modules/item';

export default class List {
	constructor(parent, elementUrl) {
		this._elementUrl = elementUrl;
		this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
		this._elementUrls = this._elementUrlIsArray ? this._elementUrl : [this._elementUrl || ''];
		this._parent = parent;
		this._contextUrlIsArray = this._parent._contextUrlIsArray;
		this._contextUrls = this._parent._contextUrls;
	}

	// Inteface

	column(elementUrl) { return new Column(this, elementUrl) }
	folder(elementUrl) { return new Folder(this, elementUrl) }
	file(elementUrl) { return new File(this, elementUrl) }
	item(elementUrl) { return new Item(this, elementUrl) }

	async get(opts) {
		return this._execute(null, spObject => (spObject.cachePath = spObject.getEnumerator ? 'properties' : 'property', spObject), opts)
	}

	async create(opts) {
		return await this._execute('create', (spContextObject, element) => {
			const listCreationInfo = new SP.ListCreationInformation;
			utility.setFields(listCreationInfo, {
				set_title: element.Title,
				set_templateType: element.BaseTemplate || SP.ListTemplateType[element.TemplateType || 'genericList'],
				set_url: element.Url,
				set_templateFeatureId: element.TemplateFeatureId,
				set_customSchemaXml: element.CustomSchemaXml,
				set_dataSourceProperties: element.DataSourceProperties,
				set_documentTemplateType: element.DocumentTemplateType,
				set_quickLaunchOption: element.QuickLaunchOption
			})
			const spObject = spContextObject.add(listCreationInfo);
			utility.setFields(spObject, {
				set_contentTypesEnabled: element.ContentTypesEnabled,
				set_defaultContentApprovalWorkflowId: element.DefaultContentApprovalWorkflowId,
				set_defaultDisplayFormUrl: element.DefaultDisplayFormUrl,
				set_defaultEditFormUrl: element.DefaultEditFormUrl,
				set_defaultNewFormUrl: element.DefaultNewFormUrl,
				set_description: element.Description,
				set_direction: element.Direction,
				set_draftVersionVisibility: element.DraftVersionVisibility,
				set_enableAttachments: element.EnableAttachments || false,
				set_enableFolderCreation: element.EnableFolderCreation === void 0 ? true : element.EnableFolderCreation,
				set_enableMinorVersions: element.EnableMinorVersions,
				set_enableModeration: element.EnableModeration,
				set_enableVersioning: element.EnableVersioning === void 0 ? true : element.EnableVersioning,
				set_forceCheckout: element.ForceCheckout,
				set_hidden: element.Hidden,
				set_imageUrl: element.ImageUrl,
				set_irmEnabled: element.IrmEnabled,
				set_irmExpire: element.IrmExpire,
				set_irmReject: element.IrmReject,
				set_isApplicationList: element.IsApplicationList,
				set_lastItemModifiedDate: element.LastItemModifiedDate,
				set_majorVersionLimit: element.EnableVersioning ? element.MajorVersionLimit : void 0,
				set_multipleDataList: element.MultipleDataList,
				set_noCrawl: element.NoCrawl === void 0 ? true : element.NoCrawl,
				set_objectVersion: element.ObjectVersion,
				set_onQuickLaunch: element.OnQuickLaunch,
				set_validationFormula: element.ValidationFormula,
				set_validationMessage: element.ValidationMessage
			});

			if (!element.BaseTemplate && !utility.FILE_LIST_TEMPLATES[element.TemplateType]) {
				utility.setFields(spObject, {
					set_documentTemplateUrl: element.DocumentTemplateUrl,
					MajorWithMinorVersionsLimit: element.EnableVersioning ? element.MajorWithMinorVersionsLimit : void 0
				})
			}
			spObject.update();
			spObject.cachePath = 'property';
			return spObject
		}, opts);
	}

	async update(opts) {
		return this._execute('update', (spObject, element) => {
			utility.setFields(spObject, {
				set_title: element.Title,
				set_enableFolderCreation: element.EnableFolderCreation,
				set_contentTypesEnabled: element.ContentTypesEnabled,
				set_defaultContentApprovalWorkflowId: element.DefaultContentApprovalWorkflowId,
				set_defaultDisplayFormUrl: element.DefaultDisplayFormUrl,
				set_defaultEditFormUrl: element.DefaultEditFormUrl,
				set_defaultNewFormUrl: element.DefaultNewFormUrl,
				set_description: element.Description,
				set_direction: element.Direction,
				set_draftVersionVisibility: element.DraftVersionVisibility,
				set_documentTemplateUrl: element.DocumentTemplateUrl,
				set_enableAttachments: element.EnableAttachments,
				set_enableMinorVersions: element.EnableMinorVersions,
				set_enableModeration: element.EnableModeration,
				set_enableVersioning: element.EnableVersioning,
				set_forceCheckout: element.ForceCheckout,
				set_hidden: element.Hidden,
				set_imageUrl: element.ImageUrl,
				set_irmEnabled: element.IrmEnabled,
				set_irmExpire: element.IrmExpire,
				set_irmReject: element.IrmReject,
				set_isApplicationList: element.IsApplicationList,
				set_lastItemModifiedDate: element.LastItemModifiedDate,
				set_majorVersionLimit: element.MajorVersionLimit,
				set_majorWithMinorVersionsLimit: element.MajorWithMinorVersionsLimit,
				set_multipleDataList: element.MultipleDataList,
				set_noCrawl: element.NoCrawl,
				set_objectVersion: element.ObjectVersion,
				set_onQuickLaunch: element.OnQuickLaunch,
				set_validationFormula: element.ValidationFormula,
				set_validationMessage: element.ValidationMessage
			});
			spObject.update();
			spObject.cachePath = 'property';
			return spObject;
		}, opts);
	}

	async delete(opts = {}) {
		return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject =>
			(spObject.recycle && spObject[opts.noRecycle ? 'deleteObject' : 'recycle'](), spObject.cachePath = 'property', spObject), opts)
	}

	async cloneLayout() {
		console.log('cloning layout in progress...');
		const elements = await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._elementUrls.map(async elementUrl => {
				const { Title, to } = elementUrl;
				const targetWebUrl = to.webUrl;
				const targetListUrl = to.listUrl;

				if (!targetWebUrl) throw new Error('Target webUrl is missed');
				if (!targetListUrl) throw new Error('Target listUrl is missed');
				if (!Title) throw new Error('Source list Title is missed');

				const targetTitle = to.Title || targetListUrl.split('/').slice(-1);

				const targetSPX = spx(targetWebUrl);
				const sourceSPX = spx(contextUrl);
				const targetSPXList = targetSPX.list(targetListUrl);
				const sourceSPXList = sourceSPX.list(Title);

				await targetSPXList.get({ silent: true }).catch(async () => {
					const newListData = Object.assign({}, await sourceSPXList.get());
					newListData.Title = targetTitle;
					delete newListData.Url;
					await targetSPX.list(newListData).create()
				});
				const [sourceColumns, targetColumns] = await Promise.all([
					sourceSPXList.column().get(),
					targetSPXList.column().get({ groupBy: 'StaticName' })
				])
				const columnsToCreate = sourceColumns.reduce((acc, sourceColumn) => {
					const targetColumn = targetColumns[sourceColumn.StaticName];
					if (!targetColumn && !sourceColumn.FromBaseType) acc.push(Object.assign({}, sourceColumn));
					return acc
				}, []);
				return columnsToCreate.length ? targetSPXList.column(columnsToCreate).create() : void 0;
			})), []))
		console.log('cloning layout is complete!');
		return elements.length > 1 ? elements : elements[0];
	}

	async clone(opts = {}) {
		const columnsToExclude = {
			Attachments: true,
			MetaInfo: true,
			FileLeafRef: true,
			Order: true
		}
		console.log('cloning in progress...');
		await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._elementUrls.map(async elementUrl => {
				const foldersToCreate = [];
				const { Title, to = {} } = elementUrl;
				const targetWebUrl = to.webUrl;
				const targetListUrl = to.listUrl;

				if (!targetWebUrl) throw new Error('Target webUrl is missed');
				if (!targetListUrl) throw new Error('Target listUrl is missed');
				if (!Title) throw new Error('Source list Title is missed');

				const targetTitle = to.Title || targetListUrl.split('/').slice(-1);

				const targetSPX = spx(targetWebUrl);
				const sourceSPX = spx(contextUrl);
				const targetSPXList = targetSPX.list(targetTitle);
				const sourceSPXList = sourceSPX.list(Title);

				await sourceSPX.list(elementUrl).cloneLayout();

				const [sourceListData, sourceColumnsData, sourceFoldersData, sourceItemsData] = await Promise.all([
					sourceSPXList.get(),
					sourceSPXList.column().get({ groupBy: 'InternalName' }),
					sourceSPXList.item({ query: 'FSObjType eq 1', scope: 'all' }).get(),
					sourceSPXList.item({ query: '', scope: 'allItems' }).get()
				])

				const foldersMapped = sourceFoldersData.map(folder => folder.FileRef.split(`${Title}/`)[1]);

				for (let folder of sourceFoldersData) {
					const props = {};
					for (let prop in folder) {
						const folderProp = folder[prop];
						if (!sourceColumnsData[prop].ReadOnlyField && !columnsToExclude[prop] && folderProp !== null) props[prop] = folderProp;
					}
					props.Url = folder.FileRef.split(`${Title}/`)[1];
					let isSubfolder;
					for (let folderUrl of foldersMapped) {
						if (new RegExp(`^${props.Url}\\/`).test(folderUrl)) {
							isSubfolder = true;
							break;
						}
					}
					!isSubfolder && foldersToCreate.push(props);
				}

				foldersToCreate.length && await targetSPXList.folder(foldersToCreate).create({ silentErrors: true }).catch(_ => _);

				if (sourceListData.BaseType) {
					for (let fileItem of sourceItemsData) {
						const fileUrl = fileItem.FileRef.split(`${Title}/`)[1];
						await sourceSPXList.file({
							Url: fileUrl,
							to: {
								webUrl: targetWebUrl,
								fileUrl: fileUrl,
								listUrl: targetListUrl
							}
						}).copy(opts)
					}
				} else {
					await targetSPXList.item(sourceItemsData.map(item => (
						item.folder = item.FileDirRef.split(`${Title}/`)[1],
						delete item.FileDirRef,
						delete item.FileLeafRef,
						delete item.FileRef,
						delete item.Order,
						item))).create();
				}
			})), []))
		console.log('cloning is complete!');
	}

	async clear(opts) {
		await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._elementUrls.map(async elementUrl =>
				this.item((await spx(contextUrl).list(elementUrl).item().get({ view: 'ID' })).map(u => u.ID)).delete(opts)
			)), []))
	}

	async getAggregations(opts = {}) {
		return await Promise.all(this._contextUrls.reduce((contextAcc, contextUrl) =>
			contextAcc.concat(this._elementUrls.map(async elementUrl => {
				let scopeStr = '';
				let fieldRefs = '';
				let caml = '';
				const aggregations = {};
				const { Title, columns, scope = 'all', query } = elementUrl;
				const clientContext = utility.getClientContext(contextUrl);
				const list = this._getSPObject(clientContext, Title);
				for (let columnName in columns) {
					fieldRefs += `<FieldRef Name="${columnName}" Type="${columns[columnName]}"/>`;
					aggregations[columnName] = 0
				};
				if (scope) {
					if (/allItems/i.test(scope)) {
						scopeStr = ' Scope="Recursive"';
					} else if (/^items$/i.test(scope)) {
						scopeStr = ' Scope="FilesOnly"';
					} else if (/^all$/i.test(scope)) {
						scopeStr = ' Scope="RecursiveAll"';
					}
				}
				if (query) caml = `<Query><Where>${getCamlQuery(query)}</Where></Query>`;
				const aggregationsQuery = list.renderListData(`<View${scopeStr}>${caml}<Aggregations>${fieldRefs}</Aggregations></View>`);
				await utility.executeQueryAsync(clientContext, opts);
				const aggregationsData = JSON.parse(aggregationsQuery.get_value()).Row[0];
				for (let name in aggregationsData) {
					const columnName = name.split('.')[0];
					if (columns.hasOwnProperty(columnName)) aggregations[columnName] = ~~aggregationsData[name];
				}
				return aggregations;
			})), []))
	}


	// Internal

	get _name() { return 'list' }

	async _execute(actionType, spObjectGetter, opts = {}) {
		const { cached } = opts;
		const elements = await Promise.all(this._contextUrls.map(async contextUrl => {
			let needToQuery;
			const clientContext = utility.getClientContext(contextUrl);
			const spObjectsToCache = new Map;
			const contextUrls = contextUrl.split('/');
			const spObjects = await Promise.all(this._elementUrls.map(elementUrl => {
				const element = this._liftElementUrlType(elementUrl);
				let listUrl = element.Title;
				if (actionType === 'create') listUrl = '';
				const spObject = spObjectGetter(this._getSPObject(clientContext, listUrl), element);
				const cachePaths = [...contextUrls, listUrl, this._name, spObject.cachePath];
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
			}));
			if (needToQuery) {
				await utility.executeQueryAsync(clientContext, opts)
				spObjectsToCache.forEach((value, key) => cache.set(key, value))
			};
			return spObjects;
		}))
		this._log(actionType, opts);
		opts.isArray = this._contextUrlIsArray || this._elementUrlIsArray;
		return utility.prepareResponseJSOM(elements, opts);
	}

	_getSPObject(clientContext, elementUrl) {
		const lists = this._parent._getSPObject(clientContext).get_lists();
		return elementUrl ? (lists[utility.isGUID(elementUrl) ? 'getById' : 'getByTitle'](elementUrl)) : lists;
	}

	_liftElementUrlType(elementUrl) {
		switch (typeOf(elementUrl)) {
			case 'object':
				if (!elementUrl.Title) elementUrl.Title = utility.getLastPath(elementUrl.Url);
				return elementUrl;
			case 'string':
				elementUrl = elementUrl.replace(/\/$/, '');
				return {
					Title: utility.getLastPath(elementUrl)
				}
		}
	}

	_log(actionType, opts = {}) {
		!opts.silent && actionType &&
			console.log(`${
				utility.ACTION_TYPES[actionType]} ${
				this._name} at ${
				this._contextUrls.join(', ')}: ${
				this._elementUrls.map(el => this._liftElementUrlType(el).Title).join(', ')} `);
	}
}