import {
	REQUEST_BUNDLE_MAX_SIZE,
	ACTION_TYPES,
	FILE_LIST_TEMPLATES,
	getInstance,
	isGUID,
	pipe,
	methodEmpty,
	method,
	methodI,
	ifThen,
	constant,
	prepareResponseJSOM,
	getClientContext,
	load,
	executorJSOM,
	getInstanceEmpty,
	setFields,
	overstep,
	hasUrlTailSlash,
	getTitleFromUrl,
	identity,
	isExists,
	isString,
	switchCase,
	typeOf,
	isArray,
	shiftSlash,
	mergeSlashes,
	arrayLast,
	isArrayFilled,
	map,
	removeEmptyUrls,
	removeDuplicatedUrls,
	AbstractBox
} from './../lib/utility';
import site from './../modules/site';
import column from './../modules/column';
import folder from './../modules/folderList';
import file from './../modules/fileList';
import item from './../modules/item';

// Internal

const NAME = 'list';

const getSPObject = elementUrl => pipe([
	getSPObjectCollection,
	ifThen(constant(isGUID(elementUrl)))([
		method('getById')(elementUrl),
		method('getByTitle')(elementUrl)
	]),
])
const getSPObjectCollection = methodEmpty('get_lists');

const liftListType = switchCase(typeOf)({
	object: list => {
		const newList = Object.assign({}, list);
		if (!list.Url) newList.Url = list.EntityTypeName || list.Title;
		if (list.Url !== '/') newList.Url = shiftSlash(newList.Url);
		if (!list.Title) newList.Title = list.EntityTypeName || list.Url;
		return newList
	},
	string: list => ({
		Url: list === '/' ? '/' : shiftSlash(mergeSlashes(list)),
		Title: getTitleFromUrl(list)
	}),
	default: _ => ({
		Url: '',
		Title: ''
	})
})


class Box extends AbstractBox {
	constructor(value = '') {
		super(value);
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([
					map(liftListType),
					removeEmptyUrls,
					removeDuplicatedUrls
				]),
				constant([liftListType()])
			])(value)
			: liftListType(value);
	}
}

const report = ({ silent, actionType, parentBox, box }) =>
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} at ${parentBox.join()}`)

const elementIterator = contextBox => elementBox => async f => {
	const clientContexts = {};
	const result = await contextBox.chain(contextElement => {
		let totalElements = 0;
		const contextUrl = contextElement.Url;
		clientContexts[contextUrl] = [getClientContext(contextUrl)];
		return elementBox.chain(element => {
			let clientContext = arrayLast(clientContexts[contextUrl]);
			if (++totalElements >= REQUEST_BUNDLE_MAX_SIZE) {
				clientContext = getClientContext(contextUrl);
				clientContexts[contextUrl].push(clientContext);
				totalElements = 0;
			}
			return f({ contextElement, clientContext, element })
		})
	});
	return { clientContexts, result }
}

// Inteface

export default (parent, urls) => {
	const instance = {
		box: getInstance(Box)(urls),
		parent,
		getSPObject,
		getSPObjectCollection
	};
	return {
		column: elements => column(instance, elements),
		folder: elements => folder(instance, elements),
		file: elements => file(instance, elements),
		item: elements => item(instance, elements),
		get: async opts => {
			const { clientContexts, result } = await elementIterator(instance.parent.box)(instance.box)(({ clientContext, element }) => {
				const parentSPObject = instance.parent.getSPObject(clientContext);
				const elementUrl = element.Url;
				const isCollection = hasUrlTailSlash(elementUrl);
				const spObject = isCollection
					? getSPObjectCollection(parentSPObject)
					: getSPObject(elementUrl)(parentSPObject);
				return load(clientContext)(spObject)(opts);
			})
			await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			return prepareResponseJSOM(opts)(result)
		},
		create: async opts => {
			const { clientContexts, result } = await elementIterator(instance.parent.box)(instance.box)(({ clientContext, element }) => {
				const parentSPObject = instance.parent.getSPObject(clientContext);
				const spObject = pipe([
					getInstanceEmpty,
					setFields({
						set_title: element.Title,
						set_templateType: element.BaseTemplate || SP.ListTemplateType[element.TemplateType || 'genericList'],
						set_url: element.Url,
						set_templateFeatureId: element.TemplateFeatureId,
						set_customSchemaXml: element.CustomSchemaXml,
						set_dataSourceProperties: element.DataSourceProperties,
						set_documentTemplateType: element.DocumentTemplateType,
						set_quickLaunchOption: element.QuickLaunchOption
					}),
					methodI('add')(getSPObjectCollection(parentSPObject)),
					overstep(setFields({
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
					})),
					overstep(ifThen(constant(!element.BaseTemplate && !FILE_LIST_TEMPLATES[element.TemplateType]))([
						setFields({
							set_documentTemplateUrl: element.DocumentTemplateUrl,
							MajorWithMinorVersionsLimit: element.EnableVersioning ? element.MajorWithMinorVersionsLimit : void 0
						})
					])),
					overstep(methodEmpty('update'))
				])(SP.ListCreationInformation)
				return load(clientContext)(spObject)(opts);
			})
			await instance.parent.box.chain(async el => {
				for (const clientContext of clientContexts[el.Url]) await executorJSOM(clientContext)(opts)
			});
			report({ ...opts, actionType: 'create', box: instance.box, parentBox: instance.parent.box });
			return prepareResponseJSOM(opts)(result)
		},
		update: async opts => {
			const { clientContexts, result } = await elementIterator(instance.parent.box)(instance.box)(({ clientContext, element }) => {
				const parentSPObject = instance.parent.getSPObject(clientContext);
				const spObject = pipe([
					setFields({
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
					}),
					overstep(methodEmpty('update'))
				])(getSPObject(element.Url)(parentSPObject))
				return load(clientContext)(spObject)(opts);
			})
			await instance.parent.box.chain(async el => {
				for (const clientContext of clientContexts[el.Url]) await executorJSOM(clientContext)(opts)
			});
			report({ ...opts, actionType: 'update', box: instance.box, parentBox: instance.parent.box });
			return prepareResponseJSOM(opts)(result)
		},
		delete: async (opts = {}) => {
			const { noRecycle } = opts;
			const { clientContexts, result } = await elementIterator(instance.parent.box)(instance.box)(({ clientContext, element }) => {
				const elementUrl = element.Url;
				const parentSPObject = instance.parent.getSPObject(clientContext);
				const spObject = getSPObject(elementUrl)(parentSPObject);
				methodEmpty(noRecycle ? 'deleteObject' : 'recycle')(spObject)
			});
			await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
			report({ ...opts, actionType: noRecycle ? 'delete' : 'recycle', box: instance.box, parentBox: instance.parent.box });
			return prepareResponseJSOM(opts)(result);
		},

		cloneLayout: async  _ => {
			console.log('cloning layout in progress...');
			await instance.parent.box.chain(async context => {
				const contextUrl = context.Url;
				return instance.box.chain(async element => {
					let targetWebUrl, targetListUrl;
					const { Title, Url, To } = element;
					if (isString(To)) {
						targetWebUrl = contextUrl;
						targetListUrl = To;
					} else {
						targetWebUrl = To.Web;
						targetListUrl = To.List;
					}
					const sourceListUrl = Title || Url;
					if (!targetWebUrl) throw new Error('Target webUrl is missed');
					if (!targetListUrl) throw new Error('Target listUrl is missed');
					if (!sourceListUrl) throw new Error('Source list Title is missed');
					const targetTitle = getTitleFromUrl(targetListUrl);

					const targetSPX = site(targetWebUrl);
					const sourceSPX = site(contextUrl);
					const targetSPXList = targetSPX.list(targetListUrl);
					const sourceSPXList = sourceSPX.list(sourceListUrl);
					await targetSPXList.get({ silent: true }).catch(async _ => {
						const newListData = Object.assign({}, await sourceSPXList.get());
						newListData.Title = targetTitle;
						await targetSPX.list(newListData).create()
					});
					const [sourceColumns, targetColumns] = await Promise.all([
						sourceSPXList.column().get(),
						targetSPXList.column().get({ groupBy: 'InternalName' })
					])
					const columnsToCreate = sourceColumns.reduce((acc, sourceColumn) => {
						const targetColumn = targetColumns[sourceColumn.InternalName];
						if (!targetColumn && !sourceColumn.FromBaseType) acc.push(Object.assign({}, sourceColumn));
						return acc
					}, []);
					return targetSPXList.column(columnsToCreate).create();
				})
			})
			console.log('cloning layout done!');
		},

		clone: async opts => {
			const columnsToExclude = {
				Attachments: true,
				MetaInfo: true,
				FileLeafRef: true,
				FileDirRef: true,
				FileRef: true,
				Order: true
			}
			console.log('cloning in progress...');
			await instance.parent.box.chain(async contextElement => {
				const contextUrl = contextElement.Url;
				return instance.box.chain(async element => {
					let targetWebUrl, targetListUrl;
					const foldersToCreate = [];
					const { Title, To } = element;
					if (isString(To)) {
						targetWebUrl = contextUrl;
						targetListUrl = To;
					} else {
						targetWebUrl = To.Web;
						targetListUrl = To.List;
					}
					const sourceListUrl = Title || Url;
					if (!targetWebUrl) throw new Error('Target webUrl is missed');
					if (!targetListUrl) throw new Error('Target listUrl is missed');
					if (!sourceListUrl) throw new Error('Source list Title is missed');

					const targetTitle = To.Title || getTitleFromUrl(targetListUrl);

					const targetSPX = site(targetWebUrl);
					const sourceSPX = site(contextUrl);
					const targetSPXList = targetSPX.list(targetTitle);
					const sourceSPXList = sourceSPX.list(sourceListUrl);

					await sourceSPX.list(element).cloneLayout();
					console.log('cloning items...');
					const [sourceListData, sourceColumnsData, sourceFoldersData, sourceItemsData] = await Promise.all([
						sourceSPXList.get(),
						sourceSPXList.column().get({ groupBy: 'InternalName', view: ['ReadOnlyField', 'InternalName'] }),
						sourceSPXList.item({ Query: 'FSObjType eq 1', Scope: 'all' }).get(),
						sourceSPXList.item({ Query: '', Scope: 'allItems' }).get()
					])
					const foldersMapped = sourceFoldersData.map(folder => folder.FileRef.split(`${sourceListUrl}/`)[1]);

					for (let folder of sourceFoldersData) {
						const props = {};
						for (let prop in folder) {
							const folderProp = folder[prop];
							if (!sourceColumnsData[prop].ReadOnlyField && !columnsToExclude[prop] && folderProp !== null) props[prop] = folderProp;
						}
						props.Url = folder.FileRef.split(`${sourceListUrl}/`)[1];
						let isSubfolder;
						for (let folderUrl of foldersMapped) {
							if (new RegExp(`^${props.Url}\\/`).test(folderUrl)) {
								isSubfolder = true;
								break;
							}
						}
						!isSubfolder && foldersToCreate.push(props);
					}
					foldersToCreate.length && await targetSPXList.folder(foldersToCreate).create({ silentErrors: true }).catch(identity);
					if (sourceListData.BaseType) {
						for (const fileItem of sourceItemsData) {
							const fileUrl = fileItem.FileRef.split(`${sourceListUrl}/`)[1];
							await sourceSPXList.file({
								Url: fileUrl,
								To: {
									Web: targetWebUrl,
									File: fileUrl,
									List: targetListUrl
								}
							}).copy(opts)
						}
					} else {
						const itemsToCreate = sourceItemsData.map(item => {
							const newItem = {};
							const folder = item.FileDirRef.split(`${sourceListUrl}/`)[1];
							if (folder) newItem.Folder = folder;
							for (const prop in item) {
								const value = item[prop];
								if (!sourceColumnsData[prop].ReadOnlyField && !columnsToExclude[prop] && isExists(value)) newItem[prop] = value;
							}
							return newItem;
						})
						return targetSPXList.item(itemsToCreate).create();
					}
				})
			})
			console.log('cloning is complete!');
		},

		clear: async opts => {
			console.log('clearing in progress...');
			await instance.parent.box.chain(context => instance.box.chain(element =>
				site(context.Url).list(element.Title).item({ Query: '' }).deleteByQuery(opts)))
			console.log('clearing is complete!');
		},

		getAggregations: opts => instance.parent.box.chain(contextElement => {
			const contextUrl = contextElement.Url;
			return instance.box.chain(async element => {
				let scopeStr = '';
				let fieldRefs = '';
				let caml = '';
				const aggregations = {};
				const { Title, Columns, Scope = 'all', Query } = element;
				const clientContext = getClientContext(contextUrl);
				const list = instance.getSPObject(Title)(instance.parent.getSPObject(clientContext));
				for (let columnName in Columns) {
					fieldRefs += `<FieldRef Name="${columnName}" Type="${Columns[columnName]}"/>`;
					aggregations[columnName] = 0
				};
				if (Scope) {
					if (/allItems/i.test(Scope)) {
						scopeStr = ' Scope="Recursive"';
					} else if (/^items$/i.test(Scope)) {
						scopeStr = ' Scope="FilesOnly"';
					} else if (/^all$/i.test(Scope)) {
						scopeStr = ' Scope="RecursiveAll"';
					}
				}
				if (Query) caml = `<Query><Where>${getCamlQuery(Query)}</Where></Query>`;
				const aggregationsQuery = list.renderListData(`<View${scopeStr}>${caml}<Aggregations>${fieldRefs}</Aggregations></View>`);
				await executorJSOM(clientContext)(opts);
				const aggregationsData = JSON.parse(aggregationsQuery.get_value()).Row[0];
				for (let name in aggregationsData) {
					const columnName = name.split('.')[0];
					if (Columns.hasOwnProperty(columnName)) aggregations[columnName] = ~~aggregationsData[name];
				}
				return aggregations;
			})
		}),

		doesUserHavePermissions: async (type = 'manageWeb') => {
			const { clientContexts, result } = await elementIterator(instance.parent.box)(instance.box)(({ clientContext, element }) => {
				const parentSPObject = instance.parent.getSPObject(clientContext);
				const spObject = getSPObject(element.Url)(parentSPObject)
				return load(clientContext)(spObject)({ view: 'EffectiveBasePermissions' });
			})
			await instance.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)())));
			return isArray(result)
				? result.map(el => el.get_effectiveBasePermissions().has(SP.PermissionKind[type]))
				: result.get_effectiveBasePermissions().has(SP.PermissionKind[type])
		}
	}
}