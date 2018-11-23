import * as utility from './../utility';
import * as cache from './../cache';
import Column from './../modules/column';
import Folder from './../modules/folderList';
import File from './../modules/file';
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
		return this._execute('create', (spContextObject, element) => {
			const listCreationInfo = new SP.ListCreationInformation;
			utility.setFields(listCreationInfo, {
				set_title: element.Title,
				set_templateType: SP.ListTemplateType[element.TemplateType || 'genericList'],
				set_url: element.Url,
				set_templateFeatureId: element.TemplateFeatureId,
				set_customSchemaXml: element.CustomSchemaXml,
				set_dataSourceProperties: element.DataSourceProperties,
				set_documentTemplateType: element.DocumentTemplateType,
				set_quickLaunchOption: element.QuickLaunchOption
			})
			const spObject = spContextObject.add(listCreationInfo);
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
			})
			spObject.update();
			spObject.cachePath = 'property';
			return spObject;
		}, opts);
	}

	async delete(opts = {}) {
		return this._execute(opts.noRecycle ? 'delete' : 'recycle', spObject =>
			(spObject.recycle && spObject[opts.noRecycle ? 'deleteObject' : 'recycle'](), spObject.cachePath = 'property', spObject), opts)
	}

	async cloneLayout(sourceListPath, targetListPath) {
		let newColumn, sourceColumnData;
		let targetColumnsData = {};
		let columnsToCreate = [];
		let sourceListData = await this.getList(sourceListPath);
		if (sourceListData) {
			let sourceListType = sourceListData.BaseType;
			let sourceColumnsData = await this.getColumnsAll(sourceListPath, {
				map: 'InternalName'
			});
			let targetListData = await this.getList(targetListPath);
			if (targetListData) {
				targetColumnsData = await this.getColumnsAll(targetListPath, {
					map: 'InternalName'
				});
				if (targetListData.ItemCount) await this.clearList(targetListPath);
			} else {
				await this.createList(targetListPath, this.LIST_TYPE_CODES[sourceListData.BaseTemplate]);
			}
			console.log('creating columns...');
			for (let columnName in sourceColumnsData) {
				sourceColumnData = sourceColumnsData[columnName];
				if (!sourceColumnData.ReadOnlyField && !sourceColumnData.FromBaseType && !targetColumnsData[columnName]) {
					newColumn = {
						title: columnName,
						type: sourceColumnData.TypeAsString,
						defaultValue: sourceColumnData.DefaultValue
					}
					switch (sourceColumnData.TypeAsString) {
						case 'Note':
							newColumn.isRichText = sourceColumnData.RichText;
							break;
						case 'Lookup':
						case 'LookupMulti':
							newColumn.lookupWebGuid = sourceColumnData.LookupWebId;
							newColumn.lookupListGuid = sourceColumnData.LookupList;
							newColumn.lookupField = sourceColumnData.LookupField;
							newColumn.isMultiple = sourceColumnData.AllowMultipleValues;
							break;
						case 'User':
						case 'UserMulti':
							newColumn.isMultiple = sourceColumnData.AllowMultipleValues;
							break;
					}
					columnsToCreate.push(newColumn);
				}
			}
			if (columnsToCreate.length) await this.createColumn(targetListPath, columnsToCreate);
			console.log('columns created');
			return [{
				list: sourceListData,
				columns: sourceColumnsData
			}, {
				list: targetListData,
				columns: targetColumnsData
			}];
		} else {
			console.log(`${sourceListData} not exists`)
		}
	}
	async clone(sourceListPath, targetListPath) {
		let sourceListType, sourceColumnData, folderData, columnValue, sourceItemData, sourceColumnName, sourceColumnValue;
		let itemsToCreate = [];
		let filesToCopy = [];
		let foldersToCreate = [];
		let columnsToCreate = {};
		let customColumnsMapped = {};
		let [source, target] = await this.cloneLayout(sourceListPath, targetListPath);
		for (let columnName in source.columns) {
			sourceColumnData = source.columns[columnName];
			if (!sourceColumnData.ReadOnlyField) customColumnsMapped[sourceColumnData.Title] = true;
		}
		if (!target.columns.sourceID) {
			await this.createColumn(targetListPath, {
				title: 'sourceID',
				type: 'Number'
			})
		}
		sourceListType = source.list.BaseType;
		console.log('get folders to create...');
		let foldersData = await this.getFoldersAll(sourceListPath, {
			needSubfolders: true,
			listType: sourceListType
		});
		console.log('folders get!');
		for (let columnName in foldersData) {
			folderData = foldersData[columnName];
			if (folderData.FolderChildCount === '0') {
				for (let j = 0; j < folderData.length; j++) {
					columnValue = folderData[j];
					if (customColumnsMapped[columnName] && (typeOf(columnValue) === 'array' && columnValue.length || columnValue !== null)) {
						columnsToCreate[columnName] = columnValue;
					}
				}
				foldersToCreate.push(folderData.FileRef);
			}
		}
		if (foldersToCreate.length) {
			console.log('creating folders...');
			await this.createFolder(targetListPath, foldersToCreate, {
				listType: sourceListType
			});
			console.log('folders created!');
		}
		console.log('get items to create...');
		let sourceItemsData = await this.getItems(sourceListPath, {
			listType: sourceListType,
			recursive: 'All',
			caml: 'FSObjType Eq 0'
		});
		console.log('items get!');
		for (let i = 0; i < sourceItemsData.length; i++) {
			sourceItemData = sourceItemsData[i];
			if (sourceItemData.File_x0020_Size) {
				filesToCopy.push(sourceItemData);
			} else if (sourceItemData.BaseType === 1) {

			} else {
				columnsToCreate = {
					sourceID: sourceItemData.ID,
					folder: extractFolderRelUrl(sourceListPath, sourceItemData.FileDirRef)
				}
				for (sourceColumnName in sourceItemData) {
					sourceColumnValue = sourceItemData[sourceColumnName];
					if (customColumnsMapped[sourceColumnName] && sourceColumnValue !== null) {
						columnsToCreate[sourceColumnName] = sourceColumnValue.get_lookupId ? sourceColumnValue.get_lookupId() : sourceColumnValue;
					}
				}
				itemsToCreate.push(columnsToCreate)
			}
		}
		this.clearCache();
		console.log('creating items...');
		await this.createItem(targetListPath, itemsToCreate);
		console.log('items created!');
		if (filesToCopy.length) {
			copyFilesR.bind(this);
			await copyFilesR(filesToCopy);
		}
		if (sourceListType === 1) {
			await this.deleteColumn(targetListPath, 'Title0');
			await this.deleteColumn(targetListPath, 'sourceID')
		}
		console.log('List cloned!');
		return await this.getList(targetListPath);

		async function copyFilesR(files) {
			if (files.length) {
				let file = files.shift();
				await this.copyFile(sourceListPath, targetListPath, file.FileLeafRef, file.FileLeafRef, {
					sourceFolder: file.FileDirRef.split('/').slice(3).join('/'),
					targetFolder: file.FileDirRef.split('/').slice(3).join('/')
				});
				return await copyFilesR(files).then(resolve);
			}
		}

		function extractFolderRelUrl(listData, relUrl) {
			let listName = listData.split('/').pop();
			return relUrl.substring(relUrl.indexOf(listName) + listName.length + 1);
		}
	}
	async merge(sourceListData, targetListData, opts = {}) {
		let mergeColumns, targetItem, sourceItem, columnName, targetColumn, targetColumnName, sourceColumnName, sourceColumn, sourceColumnItem, targetColumnItem, sourceOptColumns;
		let fromColumns = [];
		let toColumns = [];
		let {
			sourceKey = 'ID',
			targetKey = 'ID',
			columns,
			caml,
			mode = 'merge',
			mediator
		} = opts;
		if (typeOf(columns) === 'object') {
			for (let columnName in columns) {
				toColumns.push(columnName);
				if (typeOf(columns[columnName]) === 'array') {
					sourceOptColumns = columns[columnName];
					for (let i = 0; i < sourceOptColumns.length; i++) {
						fromColumns.push(sourceOptColumns[i]);
					}
				} else {
					fromColumns.push(columns[columnName]);
				}
			}
		} else {
			fromColumns = toColumns = columns;
		}
		let sourceColumns = await this.getColumn(sourceListData, fromColumns, {
			map: 'InternalName'
		});
		let targetColumns = await this.getColumn(targetListData, toColumns, {
			map: 'InternalName'
		});
		for (let targetColumnName in targetColumns) {
			let targetColumn = targetColumns[targetColumnName];
			if (typeOf(columns) === 'object') {
				if (typeOf(columns[targetColumnName]) === 'array') {
					for (let i = 0; i < columns[targetColumnName].length; i++) {
						sourceColumnName = columns[targetColumnName][i];
						sourceColumn = sourceColumns[sourceColumnName];
						if (sourceColumn.TypeAsString !== targetColumn.TypeAsString) {
							failed('"' + sourceColumnName + '" type in source is "' + sourceColumn.TypeAsString + '", in target is "' + targetColumn.TypeAsString + '"');
							return;
						}
					}
				} else {
					sourceColumnName = columns[targetColumnName];
					sourceColumn = sourceColumns[sourceColumnName];
					if (sourceColumn.TypeAsString !== targetColumn.TypeAsString) {
						failed('"' + sourceColumnName + '" type in source is "' + sourceColumn.TypeAsString + '", in target is "' + targetColumn.TypeAsString + '"');
						return;
					}
				}
			} else {
				sourceColumn = sourceColumns[targetColumnName];
				if (sourceColumn.TypeAsString !== targetColumn.TypeAsString) {
					failed('"' + targetColumnName + '" type in source is "' + sourceColumn.TypeAsString + '", in target is "' + targetColumn.TypeAsString + '"');
					return;
				}
			}
		}
		console.log('waiting for source list...');
		let sourceData = await this.getList(sourceListData, {
			expanded: true
		});
		console.log('source list get');
		let sourceListLength = sourceData.get_itemCount();
		console.log('waiting for source items...');
		let sourceItemsData = await this.getItems(sourceListData, {
			recursive: 'All',
			map: sourceKey,
		});
		console.log('source items get');
		let sourceItemsLength = Object.keys(sourceItemsData).length;
		if (sourceItemsLength) {
			console.log('waiting for target list...');
			let targetData = await this.getList(targetListData, {
				expanded: true
			});
			console.log('target list get');
			let targetListLength = targetData.get_itemCount();
			console.log('waiting for target items...');
			let targetItemsData = await this.getItems(targetListData, {
				map: targetKey,
				reqursive: 'all',
				caml: `${targetKey} isnotnull${caml ? ` and ${caml}` : ''}`
			});
			let targetItemsLength = Object.keys(targetItemsData).length;
			let mergeItems = [];
			console.log('target items get');
			console.log('source: ', sourceListLength, sourceItemsLength);
			console.log('target: ', targetListLength, targetItemsLength);
			for (let key in targetItemsData) {
				mergeColumns = {};
				targetItem = targetItemsData[key];
				sourceItem = sourceItemsData[key];
				if (!sourceItem) continue;
				if (typeOf(columns) === 'object') {
					for (let targetColumnName in columns) {
						sourceColumn = columns[targetColumnName];
						targetColumnItem = targetItem[targetColumnName];
						if (typeOf(sourceColumn) === 'array') {
							for (let i = 0; i < sourceColumn.length; i++) {
								sourceColumnName = sourceColumn[i];
								sourceColumnItem = sourceItem[sourceColumnName];
								if (sourceColumnItem !== void 0 || sourceColumnItem !== null) {
									if (mode === 'merge' && (typeOf(targetColumnItem) === 'array' ? targetColumnItem.length : targetColumnItem)) {
										// failed('"' + sourceColumnName + '" at target item "' + targetItem.ID + '" is not empty.');
										continue;
									}
									if (!sourceColumnItem) {
										// 	// failed('"' + sourceColumnName + '" at source item "' + sourceItem.ID + '" is empty.');
										continue;
									}
									mergeColumns[targetColumnName] = mediator ? mediator(sourceColumnItem) : sourceColumnItem;
									break;
								}
							}
						} else {
							for (let targetColumnName in columns) {
								sourceColumnName = columns[targetColumnName];
								sourceColumnItem = sourceItem[sourceColumnName];
								if (sourceColumnItem !== void 0 || sourceColumnItem !== null) {
									if (mode === 'merge' && (typeOf(targetColumnItem) === 'array' ? targetColumnItem.length : targetColumnItem)) {
										// failed('"' + sourceColumnName + '" at target item "' + targetItem.ID + '" is not empty.');
										continue;
									}
									if (!sourceColumnItem) {
										// 	// failed('"' + sourceColumnName + '" at source item "' + sourceItem.ID + '" is empty.');
										continue;
									}
									mergeColumns[targetColumnName] = mediator ? mediator(sourceColumnItem) : sourceColumnItem;
								}
							}
						}
					}
				} else {
					for (let i = 0; i < columns.length; i++) {
						targetColumnName = columns[i];
						sourceColumnItem = sourceItem[targetColumnName];
						targetColumnItem = targetItem[targetColumnName];
						if (sourceColumnItem !== void 0 || sourceColumnItem !== null) {
							if (mode === 'merge' && (typeOf(targetColumnItem) === 'array' ? targetColumnItem.length : targetColumnItem)) {
								// failed('"' + targetColumnName + '" at target item "' + targetItem.ID + '" is not empty.');
								continue;
							}
							if (!sourceItem[targetColumnName]) {
								// 	// failed('"' + targetColumnName + '" at source item "' + sourceItem.ID + '" is empty.');
								continue;
							}
							mergeColumns[targetColumnName] = mediator ? mediator(sourceColumnItem) : sourceColumnItem;
						}
					}
				}

				if (Object.keys(mergeColumns).length) {
					mergeColumns.ID = targetItem.ID;
					mergeItems.push(mergeColumns);
				}
			}
			if (mergeItems.length) {
				console.log(mergeItems);
				await this.updateItem(targetListData, null, mergeItems);
			} else {
				console.log('nothing to merge');
			}
		} else {
			failed('Nothing to merge.');
		}

		function failed(msg) {
			console.log('Merging failed!');
			console.log(msg);
		}
	}
	async clear(listData, noRecycle) {
		return this.item().delete({
			scope: 'all'
		});
	}
	async checkEquality(sourceListPath, targetListPath) {
		let sourceColumnData, targetColumnData, sourceColumnName, targetColumnType, sourceColumnType, sourceItemData, itemIsEqual;
		let sourceListDat = await this.getList(sourceListPath);
		let targetListData = await this.getList(targetListPath);
		console.log(sourceListData);
		console.log(targetListData);
		if (sourceListData.BaseType !== targetListData.BaseType) {
			console.log('List\'s types are not equal');
		} else {
			if (sourceListData.ItemCount !== targetListData.ItemCount) {
				console.log('Number of items are not equal');
				return
			} else {
				let sourceColumnsData = await this.getColumnsAll(sourceListPath);
				let targetColumnsData = await this.getColumnsAll(targetListPath, {
					map: 'InternalName'
				});
				console.log(sourceColumnsData);
				console.log(targetColumnsData);
				for (let i = 0; i < sourceColumnsData.length; i++) {
					sourceColumnData = sourceColumnsData[i];
					sourceColumnType = sourceColumnData.TypeAsString;
					sourceColumnName = sourceColumnData.InternalName;
					targetColumnData = targetColumnsData[sourceColumnName];
					if (targetColumnData) {
						targetColumnType = targetColumnData.TypeAsString;
						if (targetColumnType !== sourceColumnType) {
							console.log('Type of column ' + sourceColumnName + ' does not match: ' + targetColumnType + ' and ' + sourceColumnType);
							return
						}
					} else {
						console.log('Column ' + sourceColumnName + ' is missed');
						return
					}
				}
				let sourceItemsData = await this.getItems(sourceListPath);
				let targetItemsData = await this.getItems(targetListPath);
				for (let i = 0; i < sourceItemsData.length; i++) {
					sourceItemData = sourceItemsData[i];
					itemIsEqual = false;
					for (let j = 0; j < targetItemsData.length; j++) {
						targetItemData = targetItemsData[j];
						if (checkEquality(sourceItemData, targetItemData)) itemIsEqual = true;
					}
					if (!itemIsEqual) {
						console.log(`Item ${sourceItemData.ID} is not equal to any item`);
						console.log(sourceItemData);
					}
				}
				console.log('Lists are equal');

				function checkEquality(sourceItem, targetItem) {
					let columnName, sourceColumnValue, targetColumnValue, sourceLookupId, lookupEquality;
					let skipColumns = {
						Order: true
					}
					for (columnName in sourceItem) {
						sourceColumnValue = sourceItem[columnName];
						targetColumnValue = targetItem[columnName];
						if (!targetColumnsData[columnName].ReadOnlyField && !skipColumns[columnName]) {
							switch (targetColumnsData[columnName].TypeAsString) {
								case 'Text':
								case 'Note':
								case 'Number':
								case 'Boolean':
									if (sourceColumnValue !== targetColumnValue) {
										return false
									}
									break;
								case 'Lookup':
								case 'User':
									if (sourceColumnValue && targetColumnValue) {
										sourceColumnValue.get_lookupId() !== targetColumnValue.get_lookupId();
										console.log(columnName);
										return false;
									} else {
										if (sourceColumnValue || targetColumnValue) {
											return false;
										}
									}
									break;
								case 'LookupMulti':
								case 'UserMulti':
									if (sourceColumnValue.length && targetColumnValue.length) {
										if (sourceColumnValue.length === targetColumnValue.length) {
											lookupEquality = 0;
											for (let i = 0; i < sourceColumnValue.length; i++) {
												sourceLookupId = sourceColumnValue[i].get_lookupId();
												for (let j = 0; j < targetColumnValue.length; j++) {
													if (sourceLookupId === targetColumnValue[j].get_lookupId()) {
														lookupEquality++;
													}
												}
											}
											if (sourceColumnValue.length !== lookupEquality) {
												return false
											}
										} else {
											return false
										}
									} else {
										if (sourceColumnValue.length || targetColumnValue.length) {
											return false
										}
									}
									break;
							}
						}
					}
					return true;
				}
			}
		}
	}
	async getAggregations(listData, columns, opts = {}) {
		let column, typeOfColumn;
		let columnsCaml = '<IsNotNull><FieldRef Name="ID"/></IsNotNull>';
		let columnCaml = '';
		let folder = opts.folder;
		let caml = opts.caml;
		folder = folder ? `<Eq><FieldRef Name="FileDirRef"></FieldRef><Value Type="Lookup">${this.formatFolderRelUrl(listData, folder, opts.listType)}</Value></Eq>` : '';
		let fieldRefs = '';
		if (typeOf(columns) !== 'object') return;
		typeOf(listData) === 'string' && (istData = this.getListData(listData));
		for (let title in columns) {
			column = columns[title];
			typeOfColumn = typeOf(column);
			if (typeOfColumn === 'string') {
				fieldRefs += `<FieldRef Name="${title}" Type="${column}"/>`
			} else if (typeOfColumn === 'object') {
				fieldRefs += `<FieldRef Name="${title}" Type="${column.type}"/>`;
				if (column.caml) {
					columnsCaml = '<IsNull><FieldRef Name="ID"/></IsNull>';
					columnCaml = typeOf(column.caml) === 'object' ? this.getCamlQuery(column.caml) : this.parseCamlString(column.caml);
					columnsCaml = `<Or>${columnsCaml + columnCaml}</Or>`
				}
			}
		}
		let camlQuery = folder && caml ? `<And>${folder + caml}</And>` : (folder ? folder : caml || '<IsNotNull><FieldRef Name="ID"/></IsNotNull>');
		if (columnsCaml) camlQuery = `<And>${camlQuery + columnsCaml}</And>`;
		let clientContext = new SP.ClientContext(listData.path);
		let list = this.getListSPObject(clientContext, listData.title);
		let camlResult = `
				<View Scope="RecursiveAll">
					<Query><Where>${camlQuery}</Where></Query>
					<Aggregations>${fieldRefs}</Aggregations>
				</View>`
		let aggregationsQuery = list.renderListData(camlResult);
		await this.execute(clientContext);
		let aggregationsData = JSON.parse(aggregationsQuery.get_value()).Row[0];
		let aggregations = {};
		for (let columnName in columns) aggregations[columnName] = 0;
		for (let name in aggregationsData) {
			splittedName = name.split('.')[0];
			if (columns.hasOwnProperty(splittedName)) aggregations[splittedName] = ~~aggregationsData[name];
		}
		return aggregations;
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
				let listUrl = element.Url;
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
				if (!elementUrl.Url) elementUrl.Url = elementUrl.Title;
				if (!elementUrl.Title) elementUrl.Title = utility.getLastPath(elementUrl.Url);
				return elementUrl;
			case 'string':
				elementUrl = elementUrl.replace(/\/$/, '');
				return {
					Url: elementUrl,
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