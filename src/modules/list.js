import {
	ACTION_TYPES_TO_UNSET,
	ACTION_TYPES,
	FILE_LIST_TEMPLATES,
	Box,
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
	urlSplit,
	load,
	executorJSOM,
	getInstanceEmpty,
	setFields,
	overstep,
	hasUrlTailSlash,
	isFilled,
	prop,
	getTitleFromUrl,
	identity,
	isExists,
	isStringEmpty,
	slice,
	executeJSOM,
	isString
} from './../lib/utility';
import * as cache from './../lib/cache';
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

const report = ({ silent, actionType }) => parentBox => box => spObjects => (
	!silent && actionType && console.log(`${ACTION_TYPES[actionType]} ${NAME}: ${box.join()} at ${parentBox.join()}`),
	spObjects
)

const execute = parent => box => cacheLeaf => actionType => spObjectGetter => (opts = {}) => parent.box
	.chainAsync(async contextElement => {
		const { cached } = opts;
		let needToQuery = true;
		const spObjectsToCache = new Map;
		const contextUrl = contextElement.Url;
		const clientContext = getClientContext(contextUrl);
		const contextUrls = urlSplit(contextUrl);
		const parentSPObject = parent.getSPObject(clientContext);
		const spObjects = box.chain(element => {
			let listUrl = element.Title;
			const spObject = spObjectGetter({
				spContextObject: actionType === 'create' || hasUrlTailSlash(listUrl) ? getSPObjectCollection(parentSPObject) : getSPObject(listUrl)(parentSPObject),
				element
			});
			const cachePath = [...contextUrls, NAME, isStringEmpty(listUrl) ? cacheLeaf + 'Collection' : cacheLeaf, listUrl];
			ACTION_TYPES_TO_UNSET[actionType] && cache.unset(slice(0, -3)(cachePath));
			const spObjectCached = cached ? cache.get(cachePath) : null;
			if (actionType === 'delete' || actionType === 'recycle') return;
			if (cached && isFilled(spObjectCached)) {
				needToQuery = false;
				return spObjectCached;
			} else {
				const currentSPObjects = load(clientContext)(spObject)(opts);
				spObjectsToCache.set(cachePath, currentSPObjects)
				return currentSPObjects;
			}
		})
		if (needToQuery) {
			await executorJSOM(clientContext)(opts)
			spObjectsToCache.forEach((value, key) => cache.set(value)(key))
		};
		return spObjects;
	})
	.then(report({ ...opts, actionType })(parent.box)(box))
	.then(prepareResponseJSOM(opts))

// Inteface

export default (parent, urls) => {
	const instance = {
		box: getInstance(Box)(urls, 'list'),
		parent,
		getSPObject,
		getSPObjectCollection
	};
	const executeBinded = execute(parent)(instance.box)('properties');
	return {
		column: (instance => elements => column(instance, elements))(instance),
		folder: (instance => elements => folder(instance, elements))(instance),
		file: (instance => elements => file(instance, elements))(instance),
		item: (instance => elements => item(instance, elements))(instance),
		get: executeBinded(null)(prop('spContextObject')),
		create: executeBinded('create')(({ spContextObject, element }) => pipe([
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
			methodI('add')(spContextObject),
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
		])(SP.ListCreationInformation)),
		update: executeBinded('update')(({ spContextObject, element }) => pipe([
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
		])(spContextObject)),
		delete: (opts = {}) => executeBinded(opts.noRecycle ? 'delete' : 'recycle')(({ spContextObject }) =>
			overstep(methodEmpty(opts.noRecycle ? 'deleteObject' : 'recycle'))(spContextObject))(opts),

		cloneLayout: (instance => async  _ => {
			console.log('cloning layout in progress...');
			await instance.parent.box.chainAsync(async context => {
				const contextUrl = context.Url;
				return instance.box.chainAsync(async element => {
					let targetWebUrl, targetListUrl;
					const { Title, To } = element;
					if (isString(To)) {
						targetWebUrl = contextUrl;
						targetListUrl = To;
					} else {
						targetWebUrl = To.Web;
						targetListUrl = To.List;
					}
					if (!targetWebUrl) throw new Error('Target webUrl is missed');
					if (!targetListUrl) throw new Error('Target listUrl is missed');
					if (!Title) throw new Error('Source list Title is missed');
					const targetTitle = getTitleFromUrl(targetListUrl);

					const targetSPX = site(targetWebUrl);
					const sourceSPX = site(contextUrl);
					const targetSPXList = targetSPX.list(targetListUrl);
					const sourceSPXList = sourceSPX.list(Title);

					await targetSPXList.get({ silent: true }).catch(async () => {
						const newListData = Object.assign({}, await sourceSPXList.get());
						newListData.Title = targetTitle;
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
					return targetSPXList.column(columnsToCreate).create();
				})
			})
			console.log('cloning layout done!');
		})(instance),

		clone: (instance => async (opts = {}) => {
			const columnsToExclude = {
				Attachments: true,
				MetaInfo: true,
				FileLeafRef: true,
				FileDirRef: true,
				FileRef: true,
				Order: true
			}
			console.log('cloning in progress...');
			await instance.parent.box.chainAsync(async contextElement => {
				const contextUrl = contextElement.Url;
				return instance.box.chainAsync(async element => {
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

					if (!targetWebUrl) throw new Error('Target webUrl is missed');
					if (!targetListUrl) throw new Error('Target listUrl is missed');
					if (!Title) throw new Error('Source list Title is missed');

					const targetTitle = To.Title || getTitleFromUrl(targetListUrl);

					const targetSPX = site(targetWebUrl);
					const sourceSPX = site(contextUrl);
					const targetSPXList = targetSPX.list(targetTitle);
					const sourceSPXList = sourceSPX.list(Title);

					await sourceSPX.list(element).cloneLayout();
					console.log('cloning items...');
					const [sourceListData, sourceColumnsData, sourceFoldersData, sourceItemsData] = await Promise.all([
						sourceSPXList.get(),
						sourceSPXList.column().get({ groupBy: 'InternalName', view: ['ReadOnlyField', 'InternalName'] }),
						sourceSPXList.item({ Query: 'FSObjType eq 1', Scope: 'all' }).get(),
						sourceSPXList.item({ Query: '', Scope: 'allItems' }).get()
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
					foldersToCreate.length && await targetSPXList.folder(foldersToCreate).create({ silentErrors: true }).catch(identity);
					if (sourceListData.BaseType) {
						for (const fileItem of sourceItemsData) {
							const fileUrl = fileItem.FileRef.split(`${Title}/`)[1];
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
							const folder = item.FileDirRef.split(`${Title}/`)[1];
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
		})(instance),

		clear: (instance => async opts => {
			console.log('clearing in progress...');
			await instance.parent.box.chainAsync(context =>
				instance.box.chainAsync(element =>
					site(context.Url).list(element.Title).item({ Query: '' }).deleteByQuery(opts)))
			console.log('clearing is complete!');
		})(instance),

		getAggregations: (instance => opts =>
			instance.parent.box.chainAsync(contextElement => {
				const contextUrl = contextElement.Url;
				return instance.box.chainAsync(async element => {
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
			}))(instance),

		doesUserHavePermissions: (instance => (type = 'manageWeb') =>
			instance.parent.box.chainAsync(context => instance.box.chainAsync(async element => {
				const clientContext = getClientContext(context.Url);
				const web = clientContext.get_web();
				const oList = web.get_lists().getByTitle(element.Url);
				await executeJSOM(clientContext)(oList)({ view: 'EffectiveBasePermissions' });
				return oList.get_effectiveBasePermissions().has(SP.PermissionKind[type])
			})))(instance),
	}
}