//	================================================================================
//	======       ===       =====    ========    ==        ====     ===        ======
//	======  ====  ==  ====  ===  ==  ========  ===  =========  ===  =====  =========
//	======  ====  ==  ====  ==  ====  =======  ===  ========  ===========  =========
//	======  ====  ==  ===   ==  ====  =======  ===  ========  ===========  =========
//	======       ===      ====  ====  =======  ===      ====  ===========  =========
//	======  ========  ====  ==  ====  =======  ===  ========  ===========  =========
//	======  ========  ====  ==  ====  ==  ===  ===  ========  ===========  =========
//	======  ========  ====  ===  ==  ===  ===  ===  =========  ===  =====  =========
//	======  ========  ====  ====    =====     ====        ====     ======  =========
//	================================================================================

async createWebProject(webPath) {
	let index = 1;
	if (/^\//.test(webPath)) webPath = webPath.substring(1);
	let splits = webPath.split('/');
	createWeb = createWeb.bind(this);
	copyFile = copyFile.bind(this);
	if (splits.length && splits[0]) {
		createWebR = createWebR.bind(this);
		return await createWebR()
	} else {
		return await createWeb(webPath)
	}

	async function createWebR() {
		if (splits.length >= index) {
			let path = splits.slice(0, index++).join('/');
			let web = await this.getWeb(`/apps/${path}`);
			if (!web) await createWeb(path);
			return await createWebR();
		} else {
			console.log('created all webs');
		}
	}

	async function createWeb(path) {
		path = '/apps/' + path;
		await this.createWeb(path);
		let promiseLists = [];
		let listsToDelete = ['/Site Pages', '/Site Assets', '/Documents']

		for (let i = 0; i < listsToDelete.length; i++) {
			promiseLists.push(this.deleteList((path + listsToDelete[i])))
		}

		Promise.all(promiseLists).then(async () => {
			let srcListPath = path + '/src';
			await this.createList(srcListPath, {
				listType: 'documentLibrary',
				enableVersioning: true,
				description: 'Main assets list'
			});
			await this.updateList(srcListPath, {
				enableVersioning: true
			});
			let promiseFiles = [
				copyFile(path, 'aspx', null, 'default.aspx'),
				copyFile(path, 'aspx', null, 'index.aspx'),
				copyFile(path, 'js', 'src', 'index.js', 'js'),
				copyFile(path, 'css', 'src', 'style.css', 'css')
			];

			Promise.all(promiseFiles).then(() => {
				let filesToDelete = ['newsfeed.aspx', 'GettingStarted.aspx'];
				let promiseDeletes = [];
				for (let i = 0; i < filesToDelete.length; i++) {
					promiseDeletes.push(this.deleteFileByUrl(path, `${path}/${filesToDelete[i]}`))
				}
				Promise.all(promiseDeletes).then(() => {
					let foldersToDelete = ['m', 'images', '_private'];
					let promiseFolders = [];
					for (let i = 0; i < foldersToDelete.length; i++) {
						promiseFolders.push(this.deleteFolderByUrl(path, foldersToDelete[i]))
					}
					Promise.all(promiseFolders).then(() => {
						console.log('created web: ' + path);
					})
				})
			})
		})
	}

	async function copyFile(path, sourceFolder, listName, filename, targetFolder) {
		return await this.uploadFile(`${path}/${listName}`, {
			base64: btoa(await this.getFileREST('/common/ProjectLayoutAssets', filename, {
				needValue: true,
				folder: sourceFolder
			})),
			filename: filename
		}, {
				folder: targetFolder
			});
	}
}



//	=========================================================================================================
//	======       ===    ===      =====     ===  ====  ===      ====      ===    ====    ====  =======  ======
//	======  ====  ===  ===  ====  ===  ===  ==  ====  ==  ====  ==  ====  ===  ====  ==  ===   ======  ======
//	======  ====  ===  ===  ====  ==  ========  ====  ==  ====  ==  ====  ===  ===  ====  ==    =====  ======
//	======  ====  ===  ====  =======  ========  ====  ===  ========  ========  ===  ====  ==  ==  ===  ======
//	======  ====  ===  ======  =====  ========  ====  =====  ========  ======  ===  ====  ==  ===  ==  ======
//	======  ====  ===  ========  ===  ========  ====  =======  ========  ====  ===  ====  ==  ====  =  ======
//	======  ====  ===  ===  ====  ==  ========  ====  ==  ====  ==  ====  ===  ===  ====  ==  =====    ======
//	======  ====  ===  ===  ====  ===  ===  ==   ==   ==  ====  ==  ====  ===  ====  ==  ===  ======   ======
//	======       ===    ===      =====     ====      ====      ====      ===    ====    ====  =======  ======
//	=========================================================================================================

async createDiscussion(listData, columns) {
	let list = await this.getList(listData, {
		expanded: true
	})
	let clientContext = list.get_context();
	let discussionItem = SP.Utilities.Utility.createNewDiscussion(clientContext, list, columns.Subject);
	for (let columnName in columns) {
		if (columnName === 'Subject') continue;
		discussionItem.set_item(columnName, columns[columnName]);
	}
	discussionItem.update();
	clientContext.load(discussionItem)
	await this.execute(clientContext);
	return [discussionItem.get_id(), discussionItem];
}

async createReply(listData, id, columns) {
	let item = await this.getItemById(listData, id, {
		expanded: true
	})
	let clientContext = item.get_context();
	let replyItem = SP.Utilities.Utility.createNewDiscussionReply(clientContext, item);
	for (let columnName in columns) {
		replyItem.set_item(columnName, columns[columnName]);
	}
	replyItem.update();
	clientContext.load(replyItem);
	await this.execute(clientContext);
	return [replyItem.get_id(), replyItem];
}


//	========================================
//	======        =====  ======      =======
//	=========  =======    ====   ==   ======
//	=========  ======  ==  ===  ====  ======
//	=========  =====  ====  ==  ============
//	=========  =====  ====  ==  ============
//	=========  =====        ==  ===   ======
//	=========  =====  ====  ==  ====  ======
//	=========  =====  ====  ==   ==   ======
//	=========  =====  ====  ===      =======
//	========================================

async getTagRefItemByRefIdAndRefListData(refListData, refId) {
	return await this.getItemById(refListData, refId)
}

async getTagRefItemByIdAndRefListData(refListData, id) {
	let itemsData = await this.getItems(refListData, {
		recursive: true,
		caml: `Lookup ${this.TAGS_COLUMN_NAMES[refListData]} Includes ${id}`
	})
	return itemsData[0];
}

async getTagRefItemsByIdAndRefListData(refListData, ids) {
	if (ids && ids[0] && ids[0].get_lookupId) {
		return ids.map(item => {
			return item.get_lookupId()
		})
	} else {
		return await this.getItems(refListData, {
			recursive: true,
			caml: `Lookup ${this.TAGS_COLUMN_NAMES[refListData]} Includes ${ids}`
		})
	}
}

async getTagRefItemByRefIdAndId(refListData, id, refId) {
	let itemsData = await this.getItems(refListData, {
		recursive: true,
		caml: `Lookup ${this.TAGS_COLUMN_NAMES[refListData]} Includes ${id} and ID Eq ${refId}`
	})
	return itemsData[0];
}

async getTagRefItemById(id) {
	return await this.getTagRefItemByTagEntry(await this.getTagEntryById(id));
}

async getTagRefItemByIdIncludes(id) {
	let itemDataToUpdate;
	let promises = [];
	for (let listData in this.TAGS_COLUMN_NAMES) {
		promises.push(new Promise(async (resolve, reject) => {
			let itemData = await this.getTagRefItemByIdAndRefListData(listData, id);
			if (itemData) {
				itemDataToUpdate = {
					refId: itemData.ID,
					refListData: listData
				}
			}
			resolve();
		}))
	}
	await Promise.all(promises)
	return itemDataToUpdate;
}

async getTagListDataByRefIdAndIdIncludes(id, refId) {
	let deferred, foundListData;
	let promises = [];
	for (let listData in this.TAGS_COLUMN_NAMES) {
		promises.push(new Promise(async (resolve, reject) => {
			let itemData = await this.getTagRefItemByRefIdAndId(listData, id, refId)
			if (itemData) foundListData = listData;
			resolve();
		}))
	}
	await Promise.all(promises)
	return foundListData;
}

async getTagRefItemByTagEntry(tagData) {
	let deferred, foundListData, itemDataToUpdate;
	let deferreds = [];
	let tagEntriesToUpdate = [];
	let tagEntriesToDelete = [];
	if (tagData.refId) {
		if (tagData.refListData) {
			let itemData = await this.getTagRefItemByRefIdAndRefListData(tagData.refListData, tagData.refId);
			if (!itemData || !this.checkTagRefItemById(tagData.ID, itemData[this.TAGS_COLUMN_NAMES[tagData.refListData]])) {
				return {
					ID: tagData.ID
				}
			}
		} else {
			let listData = await this.getTagListDataByRefIdAndIdIncludes(tagData.ID, tagData.refId);
			let returnObj = {
				ID: tagData.ID,
			}
			if (listData) returnObj.refListData = listData;
			return returnObj;
		}
	} else {
		let returnObj = {
			ID: tagData.ID,
		}
		if (tagData.refListData) {
			let imtemData = await this.getTagRefItemByIdAndRefListData(tagData.refListData, tagData.ID)
			if (itemData && this.checkTagRefItemById(tagData.ID, itemData[this.TAGS_COLUMN_NAMES[tagData.refListData]])) returnObj.refId = itmemData.ID;
			return returnObj;
		} else {
			let tagEntriesToUpdate = await this.getTagRefItemByIdIncludes(tagData.ID)
			if (tagEntriesToUpdate) returnObj.columns = tagEntriesToUpdate
			return returnObj;
		}
	}
}

async getTagEntryById(id) {
	return await this.getItemById(this.LIBRARY.tags, id);
}

async getTagEntryDuplicates() {
	let [tagItems, tagIds] = await this.getItemDuplicates(this.LIBRARY.tags, ['refId', 'refListData', 'tag']);
	return tagItems ? [tagItems, tagIds] : void 0;
}

async getTagEntryEmptyRefIds() {
	return await this.getItems(this.LIBRARY.tags, {
		view: ['ID', 'refListData', 'refId'],
		recursive: true,
		caml: 'Number refId IsNull and refListData IsNotNull'
	})
}

async getTagEntryEmptyListDatas() {
	return await this.getItems(this.LIBRARY.tags, {
		recursive: true,
		caml: 'refListData IsNull'
	})
}

async getTagEntryEmptyRefTag() {
	return await this.getItems(this.LIBRARY.tags, {
		view: ['ID', 'refListData', 'refId'],
		recursive: true,
		caml: 'Lookup tag IsNull'
	})
}

async getTagEntriesByRefId(refId, opts = {}) {
	operand = typeOf(refId) === 'array' ? 'In' : 'Eq';
	opts.recursive = true;
	opts.caml = 'Number refId ' + operand + ' ' + refId;
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTagEntriesByRefIds(refId, opts = {}) {
	operand = typeOf(refId) === 'array' ? 'In' : 'Eq';
	opts.recursive = true;
	opts.caml = `Number refId ${operand} ${refId}`;
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTagAllEntries(opts = {}) {
	opts.caml = 'FSObjType Eq 0 with LookupId=false';
	opts.recursive = true;
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTagFolder(folderTitle, opts = {}) {
	opts.caml = 'FSObjType Eq 1 and Title Eq ' + folderTitle.encode();
	let folderTagsData = await this.getItems(this.LIBRARY.tags, opts);
	return folderTagsData[0];
}

async getTagAllFolders(opts = {}) {
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTagEmptyFolders(opts = {}) {
	opts.caml = 'FSObjType Eq 1 and ItemChildCount Eq 0';
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTagRefItemsToInstall(listData, sourceColumnTitle, targetColumnTitle) {
	return await this.getItems(listData, {
		view: ['ID', sourceColumnTitle, targetColumnTitle],
		caml: `${sourceColumnTitle} IsNotNull and ${targetColumnTitle} IsNull`
	})
}

async getTagInvalidEntries(opts = {}) {
	opts.caml = '(refListData IsNull or Number refId IsNull) and FSObjType Eq 0 with LookupId=false';
	opts.recursive = true;
	return await this.getItems(this.LIBRARY.tags, opts)
}

async getTaggedLists() {
	return await this.getItems(this.LIBRARY.tags, {
		recursive: true,
		map: 'refListData'
	})
}

async checkTagEntryIntegrityById(id) {
	let tagData = await this.getTagRefItemById(id);
	if (tagData) {
		console.log(`Tag entry must be ${tagData.columns ? `restored at ${Object.keys(tagData.columns)}` : 'deleted'}`);
		return tagData
	} else {
		console.log('Tag\'s integrity is ok');
	}
}

async restoreTagEntryIntegrityById(id) {
	let tagData = await this.getTagRefItemById(id)
	if (tagData) {
		if (tagData.columns) {
			return await this.updateTagEntry(tagData.ID, tagData.columns)
		} else {
			return await this.deleteTagEntry(tagData.ID)
		}
	} else {
		console.log('Nothing to restore');
	}
}

async restoreTagEntriesIntegrityRefTags() {
	let folder, emptyTagData;
	let tagsToUpdate = [];
	let promises = [];
	let emptyTagsData = await this.getTagEntryEmptyRefTag()
	for (let i = 0; i < emptyTagsData.length; i++) {
		emptyTagData = emptyTagsData[i];
		folder = emptyTagData.FileDirRef.split('/').pop();
		promises.push(new Promise(async (resolve, reject) => {
			let tagFolderData = await this.getTagFolder(folder)
			tagsToUpdate.push({
				ID: id,
				tag: tagFolderData.ID
			})
			resolve();
		}));
	}
	await Promise.all(promises);
	console.log(tagsToUpdate);
	return await this.updateTagEntry(null, tagsToUpdate);
}

async checkTagEntriesIntegrity(opts = {}) {
	let select, caml;
	let tagsDataToUpdate = [];
	let tagsDataToDelete = [];
	select = opts.select || 'invalid';
	switch (select) {
		case 'all':
			caml = 'FSObjType Eq 0 with LookupId=false'
			break;
		case 'invalid':
		default:
			caml = '(refListData IsNull or Number refId IsNull) and FSObjType Eq 0 with LookupId=false'
			break;
	}
	let pager = new Pager(this.LIBRARY.tags, {
		pageSize: 50,
		recursive: true,
		caml: caml
	});
	for (var listData in this.TAGS_COLUMN_NAMES) {
		await getItemsByListDataR.call(this, listData);
	}
	return [tagsDataToUpdate, tagsDataToDelete];

	async function getItemsByListDataR(listData) {
		let tagsMapped = {};
		let tagsData = await pager.moveNext();
		for (let i = 0, tagData; i < tagsData.length; i++) {
			tagData = tagsData[i];
			tagsMapped[tagData.ID] = tagData;
		}
		let tagIds = Object.keys(tagsMapped);
		if (tagIds) {
			console.log(listData);
			let refItemsData = await this.getItems(listData, {
				recursive: 'All',
				caml: `Integer ${this.TAGS_COLUMN_NAMES[listData]} Includes ${tagIds} with LookupId=true`
			})
			checkTagItems.call(this, listData, tagsMapped, refItemsData)
			return await getItemsByListDataR.call(this, tagsMapped);
		} else {
			tagsDataToDelete = tagsDataToDelete.concat(Object.keys(tagsMapped).map((item) => {
				return ~~item
			}));
		}
	}

	function checkTagItems(listData, tagsMapped, refItemsData) {
		let refId;
		for (let tid in tagsMapped) {
			refId = this.checkTagRefItems(tid, this.TAGS_COLUMN_NAMES[listData], refItemsData);
			if (refId) {
				delete tagsMapped[tid];
				tagsDataToUpdate.push({
					ID: ~~tid,
					refId: refId,
					refListData: listData
				})
			}
		}
	}
}

async restoreTagEntriesIntegrity(opts) {
	let [tagEntriesToUpdate, tagEntriesToDelete] = await this.checkTagEntriesIntegrity(opts);
	if (!(tagEntriesToUpdate.length || tagEntriesToDelete.length)) {
		console.log('Tag integrity is ok');
		resolve();
		return
	}
	await Promise.all([new Promise(async (resolve, reject) => {
		if (tagEntriesToUpdate.length) {
			await this.updateTagEntry(null, tagEntriesToUpdate)
			console.log('Invalid tag entries updated');
			resolve();
		} else {
			resolve();
		}
	}), new Promise(async (resolve, reject) => {
		if (tagEntriesToDelete) {
			await this.deleteTagEntry(tagEntriesToDelete);
			console.log('Unrefed tag entries deleted');
			resolve();
		} else {
			resolve()
		}
	})])
	console.log('Tag integrity restored');
}

async createTagFolder(tagTitle) {
	return await this.createFolder(this.LIBRARY.tags, tagTitle)
}

async createTagFolders(refItemTagTitle) {
	let it = this;
	let tagFolders = await this.getItems(this.LIBRARY.tags, {
		caml: 'FSObjType Eq 1'
	})
	console.log(tagFolders);
	let tagFolder;
	let foldersMapped = [];
	for (let i = 0; i < tagFolders.length; i++) {
		tagFolder = tagFolders[i];
		foldersMapped[tagFolder.Title.toLowerCase()] = tagFolder;
	}
	let tagsData = await this.getTagRefItems();
	let tagData, tag;
	let tags = [];
	let tagsToCreateMapped = {};
	for (let i = 0; i < tagsData.length; i++) {
		tagData = tagsData[i];
		tags = JSON.parse(tagData[refItemTagTitle]);
		for (let j = 0; j < tags.length; j++) {
			tag = tags[j].toLowerCase().replace(/[^a-zа-яё0-9\-]/gi, ' ').trim();
			tag = tag.replace(/\s+/gi, ' ');
			if (tag && !foldersMapped[tag]) {
				tagsToCreateMapped[tag] = true;
			}
		}
	}
	tagsToCreate = Object.keys(tagsToCreateMapped);
	console.log(tagsToCreate.length);
	return await this.createTagFolder(tagsToCreate);
}

async createTagEntry(column) {
	return await this.createItem(this.LIBRARY.tags, column);
}

async createTagEntries(columns) {
	return await this.createItems(this.LIBRARY.tags, columns);
}

async updateTagEntry(column) {
	return await this.updateItem(this.LIBRARY.tags, column)
}

async updateTagEntries(columns) {
	return await this.updateItem(this.LIBRARY.tags, columns)
}

async deleteTagEntry(tagId) {
	return await this.deleteItem(this.LIBRARY.tags, tagId);
}

async deleteTagEntries(tagIds) {
	return await this.deleteItems(this.LIBRARY.tags, tagIds);
}

async deleteTagEntryByAuthorId(uid) {
	let entriesData = await this.getItems(this.LIBRARY.tags, {
		caml: `Author Eq ${uid}`,
		recursive: true,
		map: 'ID'
	});
	return await this.deleteItem(this.LIBRARY.tags, Object.keys(entriesData));
}

async deleteTagEntryDuplicates() {
	let [tagItems, tagIds] = await this.getTagEntryDuplicates();
	return await this.deleteTagEntry(tagIds);
}

async deleteTagEmptyFolders() {
	let tagFoldersData = await this.getTagEmptyFolders({
		map: 'ID'
	});
	return await this.deleteItem(this.LIBRARY.tags, Object.keys(tagFoldersData))
}

async deleteTagNullFolders() {
	let tagFoldersData = await this.getItems(this.LIBRARY.tags, {
		map: 'ID',
		caml: 'FSObjType Eq 1 and Title IsNull'
	});
	return await this.deleteItem(this.LIBRARY.tags, Object.keys(tagFoldersData))
}

async deleteTagEntriesByRefId(refId) {
	let tagItemsData = await this.getTagEntriesByRefId(refId, {
		map: 'ID'
	});
	return await this.deleteTagEntry(Object.keys(tagItemsData))
}

async deleteTagEntriesByRefIds(refIds) {
	let tagItemsData = await this.getTagEntriesByRefIds(refIds, {
		map: 'ID'
	});
	if (Object.keys(tagItemsData).length) {
		return await this.deleteTagEntry(Object.keys(tagItemsData), resolve)
	}
}

async installTagRefItemsByTextColumn(listData, sourceColumnTitle, targetColumnTitle) {
	createR = createR.bind(this);
	getTagsToCreateR = getTagsToCreateR.bind(this);
	let tagsData = await this.getTagRefItemsToInstall(listData, sourceColumnTitle, targetColumnTitle);
	console.log(tagsData);
	return await createR(tagsData);

	async function createR(tagsData) {
		if (tagsData.length) {
			let tagData = tagsData.shift();
			let tags = JSON.parse(tagData[sourceColumnTitle]);
			let tagsToCreate = await getTagsToCreateR(tags);
			let [ids, itemsData] = await this.createTagEntry(tagsToCreate);
			await this.updateItem(listData, tagData.ID, {
				tags: ids
			});
			return await createR(tagsData);
		}
	}

	async function getTagsToCreateR(tags) {
		if (tags.length) {
			let rawTag = tags.shift();
			// console.log(rawTag);
			let tag = rawTag.toLowerCase().replace(/[^a-zа-яё0-9\-]/gi, ' ').replace(/\s+/gi, ' ').trim();
			if (tag) {
				let tagFolderData = await this.getTagFolder(tag);
				// console.log(tagFolderData);
				let folderId = tagFolderData.ID;
				tagsToCreate.push({
					Title: tag,
					refId: tagData.ID,
					tag: folderId,
					refListData: listData,
					folder: tag
				})
				return await getTagsToCreateR(tags);
			} else {
				return await getTagsToCreateR(tags);
			}
		} else {
			return tagsToCreate;
		}
	}
}

async copyTagEntries(listData, sourceColumn, targetColumn) {
	let itemsData = await this.getItems(listData, {
		recursive: 'All',
		caml: `Lookup ${sourceColumn} IsNotNull`
	});
	let foldersData = await this.getTagAllFolders();
	let itemData, tagData, folderTitle, folderData;
	let existedFolders = {};
	for (let i = 0; i < foldersData.length; i++) {
		folderData = foldersData[i];
		existedFolders[folderData.Title.toLowerCase()] = folderData.ID;
	}
	updateTagItemR = updateTagItemR.bind(this);
	return await updateTagItemR(itemsData);

	async function updateTagItemR(itemsData) {
		let itemsToCreate = [];
		if (itemsData.length) {
			let itemData = itemsData.shift();
			for (let i = 0; i < itemData[sourceColumn].length; i++) {
				tagData = itemData[sourceColumn][i];
				folderTitle = tagData.get_lookupValue();
				itemsToCreate.push({
					Title: folderTitle,
					refId: itemData.ID,
					tag: existedFolders[folderTitle.toLowerCase()],
					refListData: listData,
					folder: folderTitle
				})
			}
			let [ids, items] = await this.createTagEntry(itemsToCreate);
			let columnToUpdate = {};
			columnToUpdate[targetColumn] = ids;
			await this.updateItem(listData, itemData.ID, columnToUpdate);
			return await updateTagItemR(itemsData).then(resolve);
		}
	}
}

async cloneTagEntries(listData, columnTitle) {
	let tempColumnName = 'tempTags';
	let columnsData = await this.getColumnsAll(listData, {
		map: 'InternalName'
	});
	if (!columnsData[tempColumnName]) {
		this.clearCache();
		await this.createTagsColumn(listData, tempColumnName)
	}
	console.log('merging columns...');
	await this.mergeColumns(listData, columnTitle, tempColumnName);
	console.log('columns merged!');
	console.log('copy refItem tag columns...');
	return await this.copyTagEntries(listData, tempColumnName, columnTitle);
}

checkTagRefItems(id, tagColumnTitle, refItemsData) {
	for (let i = 0; i < refItemsData.length; i++) {
		if (this.checkTagRefItemById(id, refItemsData[i][tagColumnTitle])) {
			return refItemsData[i].ID;
		}
	}
	return false
}

checkTagRefItemById(id, tagColumn) {
	for (let i = 0; i < tagColumn.length; i++) {
		if (tagColumn[i].get_lookupId() == id) {
			return true
		}
	}
	return false;
}

//	======================================================================
//	=======      ===        =====  =====       =====     ===  ====  ======
//	======  ====  ==  ==========    ====  ====  ===  ===  ==  ====  ======
//	======  ====  ==  =========  ==  ===  ====  ==  ========  ====  ======
//	=======  =======  ========  ====  ==  ===   ==  ========  ====  ======
//	=========  =====      ====  ====  ==      ====  ========        ======
//	===========  ===  ========        ==  ====  ==  ========  ====  ======
//	======  ====  ==  ========  ====  ==  ====  ==  ========  ====  ======
//	======  ====  ==  ========  ====  ==  ====  ===  ===  ==  ====  ======
//	=======      ===        ==  ====  ==  ====  ====     ===  ====  ======
//	======================================================================

async searchItems(value, opts = {}) {
	let clientContext = new SP.ClientContext();
	let keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);
	let searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
	let queryTemplates = ['{searchboxquery}',
		'-FileExtension:jpg',
		'-FileExtension:jpeg',
		'-FileExtension:png',
		'-Site:http://spps001*',
		'-contentclass:STS_Site',
		'-Site:http://aura.dme.aero.corp/wikilibrary/wiki/Lists*',
		'-Site:http://aura.dme.aero.corp/app*',
		'-Site:http://aura.dme.aero.corp/Tikhonchuk*',
		'-Site:http://aura.dme.aero.corp/SitePages*',
		'-Site:http://aura.dme.aero.corp/Pages/sys-*',
		'-Site:http://aura.dme.aero.corp/PublishingImages*',
		'-Site:http://aura.dme.aero.corp/Survey*',
		'-Site:http://aura.dme.aero.corp/News/event*',
		'-Site:http://aura.dme.aero.corp/social*',
		'-Site:http://aura.dme.aero.corp/News/Lists/Posts/AllPosts.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/Posts/News_Disabled.aspx',
		'-Site:http://aura.dme.aero.corp/News/Pages*',
		'-Site:http://aura.dme.aero.corp/News/SitePages*',
		'-Site:http://aura.dme.aero.corp/News/DocsI*',
		'-Site:http://aura.dme.aero.corp/News/Documents*',
		'-Site:http://aura.dme.aero.corp/News/Categories*',
		'-Site:http://aura.dme.aero.corp/News/Lists/Comments*',
		'-Site:http://aura.dme.aero.corp/News/Lists/List1*',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/AllItems.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/calendar.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/List_CurrentEvents.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/view*',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/mod-view.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/mod-view.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/MyItems.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List/my-sub.aspx',
		'-Site:http://aura.dme.aero.corp/News/Lists/List2*',
		'-Site:http://aura.dme.aero.corp/News/Lists/List3*',
		'-Site:http://aura.dme.aero.corp/News/Int*',
		'-Site:http://aura.dme.aero.corp/News/PublishingImages*',
		'-Site:http://aura.dme.aero.corp/Intellect/km*',
		'-Site:http://aura.dme.aero.corp/Intellect/OZN_Docs/*',
		'-Site:http://aura.dme.aero.corp/board/Lists/Discussions/*',
		'-Site:http://aura.dme.aero.corp/buro/Lists/MagazineIssues/*',
		'-Site:http://aura.dme.aero.corp/Intellect/%D0%9E%D1%82%D1%87%D0%B5%D1%82%D1%8B%20%D0%BE%20%D0%BA%D0%BE%D0%BC%D0%B0%D0%BD%D0%B4%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0%D1%85*',
		'-Site:http://aura.dme.aero.corp/Intellect/PublishingImages*',
		'-Site:http://aura.dme.aero.corp/Intellect/StorageLib*',
		'-Site:http://wiki.aura.dme.aero.corp/Media%20Gallery/*',
		'-Site:http://wiki.aura.dme.aero.corp/Lists/List1*',
		'-Site:http://wiki.aura.dme.aero.corp/survey*',
		'-Site:http://wiki.aura.dme.aero.corp/Pages*',
		'-Site:http://wiki.aura.dme.aero.corp/Pages/*',
		'-Site:http://wiki.aura.dme.aero.corp/*',
		'-Site:http://aura.dme.aero.corp/News*',
		'-Site:http://aura.dme.aero.corp/board*',
		'-Site:http://aura.dme.aero.corp/board/Lists/Views/*',
		'-Site:http://aura.dme.aero.corp/Intellect/Lists/Training/*',
		'-Site:http://aura.dme.aero.corp/board/Lists/DiscussionUsers/*',
		'-Site:http://wiki.aura.dme.aero.corp/Lists/Survey*',
		'-Site:http://wiki.aura.dme.aero.corp/PublishingImages*',
		'-Site:http://aura.dme.aero.corp/marketingpresentations/DocLib/Forms/*',
		'-Site:http://aura.dme.aero.corp/crowd*',
		'-Site:http://aura.dme.aero.corp/Intellect/Lists/iLikeDislikes/*',
		'-Site:http://aura.dme.aero.corp/Intellect/Lists/Discussions/*',
		'-Site:http://aura.dme.aero.corp/Intellect/Lists/DiscussionUsers/*',
		'-Site:http://aura.dme.aero.corp/AM/*',
		'-Site:http://aura.dme.aero.corp/System/*',
		'-Site:http://aura.dme.aero.corp/social/*',
		'-Site:http://aura.dme.aero.corp/Forum/*',
		'-Site:http://mysites.aura.dme.aero.corp/*',
		'-Site:http://aura.dme.aero.corp/Intellect/Lists/*',
		'-Site:http://aura.dme.aero.corp/buro/Lists/*',
		'-Site:http://aura.dme.aero.corp/buro/Lists/XEvents/*',
		'-Site:http://aura.dme.aero.corp/buro/Lists/XEventFlow/*',
		'-Site:http://aura.dme.aero.corp/buro/DocLib4/Forms/*'
	];
	if (!value) {
		value = '';
	}
	console.log(keywordQuery);
	let sortList = keywordQuery.get_sortList();
	sortList.add('[formula:rank]', 1);
	keywordQuery.set_queryText(value.toLowerCase());
	keywordQuery.set_clientType(opts.clientType || 'AllResultsQuery');
	// keywordQuery.set_sourceId('8413cd39-2156-4e00-b54d-11efd9abdb89');
	// keywordQuery.set_ignoreSafeQueryPropertiesTemplateUrl(false);
	keywordQuery.set_queryTemplate(queryTemplates.join(' '));
	keywordQuery.set_refiners(opts.refiners || 'FileType(filter=21/0/*),contentclass(filter=10/0/*),ContentTypeId(filter=15/0/*),WebTemplate(filter=10/0/*),DisplayAuthor(filter=9/0/*)');
	keywordQuery.set_rowsPerPage(opts.rowsPerPage || 10);
	keywordQuery.set_totalRowsExactMinimum(opts.set_totalRowsExactMinimum || 11);
	opts.blockDedupeMode !== void 0 && keywordQuery.set_blockDedupeMode(opts.blockDedupeMode);
	opts.bypassResultTypes !== void 0 && keywordQuery.set_bypassResultTypes(opts.bypassResultTypes);
	opts.collapseSpecification !== void 0 && keywordQuery.set_collapseSpecification(opts.collapseSpecification);
	opts.culture !== void 0 && keywordQuery.set_culture(opts.culture);
	opts.desiredSnippetLength !== void 0 && keywordQuery.set_desiredSnippetLength(opts.desiredSnippetLength);
	opts.enableInterleaving !== void 0 && keywordQuery.set_enableInterleaving(opts.enableInterleaving);
	opts.enableNicknames !== void 0 && keywordQuery.set_enableNicknames(opts.enableNicknames);
	opts.enableOrderHitHighlightedProperty !== void 0 && keywordQuery.set_enableOrderHitHighlightedProperty(opts.enableOrderHitHighlightedProperty);
	opts.enablePhonetic !== void 0 && keywordQuery.set_enablePhonetic(opts.enablePhonetic);
	opts.enableQueryRules !== void 0 && keywordQuery.set_enableQueryRules(opts.enableQueryRules);
	opts.enableSorting !== void 0 && keywordQuery.set_enableSorting(opts.enableSorting);
	opts.enableStemming !== void 0 && keywordQuery.set_enableStemming(opts.enableStemming);
	opts.generateBlockRankLog !== void 0 && keywordQuery.set_generateBlockRankLog(opts.generateBlockRankLog);
	opts.hiddenConstrains !== void 0 && keywordQuery.set_hiddenConstrains(opts.hiddenConstrains);
	opts.hitHighlightedMultivaluePropertyLimit !== void 0 && keywordQuery.set_hitHighlightedMultivaluePropertyLimit(opts.hitHighlightedMultivaluePropertyLimit);
	opts.impressionID !== void 0 && keywordQuery.set_impressionID(opts.impressionID);
	opts.maxSnippetLength !== void 0 && keywordQuery.set_maxSnippetLength(opts.maxSnippetLength);
	opts.objectVersion !== void 0 && keywordQuery.set_objectVersion(opts.objectVersion);
	opts.personalizationData !== void 0 && keywordQuery.set_personalizationData(opts.personalizationData);
	opts.processBestBets !== void 0 && keywordQuery.set_processBestBets(opts.processBestBets);
	opts.processPersonalFavorites !== void 0 && keywordQuery.set_processPersonalFavorites(opts.processPersonalFavorites);
	opts.queryTag !== void 0 && keywordQuery.set_queryTag(opts.queryTag);
	opts.rankingModelId !== void 0 && keywordQuery.set_rankingModelId(opts.rankingModelId);
	opts.refinementFilters !== void 0 && keywordQuery.set_refinementFilters(opts.refinementFilters);
	opts.reorderingRules !== void 0 && keywordQuery.set_refinementFilters(opts.refinementFilters);
	opts.resultsUrl !== void 0 && keywordQuery.set_resultsUrl(opts.resultsUrl);
	opts.rowLimit !== void 0 && keywordQuery.set_rowLimit(opts.rowLimit);
	opts.startRow !== void 0 && keywordQuery.set_startRow(opts.startRow);
	opts.showPeopleNameSuggestions !== void 0 && keywordQuery.set_showPeopleNameSuggestions(opts.showPeopleNameSuggestions);
	opts.summaryLength !== void 0 && keywordQuery.set_summaryLength(opts.summaryLength);
	opts.timeZoneId !== void 0 && keywordQuery.set_timeZoneId(opts.timeZoneId);
	opts.timeout !== void 0 && keywordQuery.set_timeout(opts.timeout);
	opts.trimDuplicates !== void 0 && keywordQuery.set_trimDuplicates(opts.trimDuplicates);
	opts.trimDuplicatesIncludeId !== void 0 && keywordQuery.set_trimDuplicatesIncludeId(opts.trimDuplicatesIncludeId);
	opts.uiLanguage !== void 0 && keywordQuery.set_uiLanguage(opts.uiLanguage);
	this.load(clientContext, keywordQuery);
	let results = searchExecutor.executeQuery(keywordQuery);
	await this.execute(clientContext);
	console.log(keywordQuery.get_objectData().get_properties());
	let resultTables = results.get_value().ResultTables;
	let unitedTables = [];
	console.log(resultTables);
	if (resultTables.length) {
		for (let i = 0; i < resultTables.length; i++) {
			unitedTables = unitedTables.concat(resultTables[i].ResultRows)
		}
		return unitedTables
	}
}