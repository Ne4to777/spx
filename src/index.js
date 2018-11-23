import $ from 'jquery';
import siteTest from './test/site'

// import * as R from "ramda";
import axios from 'axios';
import * as utility from './utility';
// import {
// 	webBundle,
// 	folderBundle,
// 	fileBundle
// } from './test'
import spx from './modules/site'
import {
	getCamlQuery,
	getCamlView,
	camlLog,
	concatQuery
} from './query_parser';
window.spx = spx;
window.log = log;
window.getCamlView = getCamlView;

// siteTest()

// let files = [];
// if (false) {
// 	for (let i = 0; i < 255; i++) {
// 		files.push({
// 			Url: `test${i}.txt`,
// 			Content: 'hi',
// 			Overwrite: true
// 		})
// 	}
// 	spx('/test/spx/testFolder').list('files').file(files).create().then(log);
// }

// $('#send').click(e => {
// 	e.preventDefault();
// 	spx('/test/spx/testFolder2').list('files').file('image-2.jpg').get({
// 		blob: true
// 	}).then(data => {
// 		console.log(data);
// 		showFile(data, 'downloaded.jpg', 'image/jpeg')
// 	});
// })


// function showFile(data, filename, mime) {
// 	// let blob = new Blob([data], {
// 	// 	type: mime || 'application/octet-stream'
// 	// })
// 	// if (window.navigator && window.navigator.msSaveOrOpenBlob) {
// 	// 	window.navigator.msSaveOrOpenBlob(blob);
// 	// 	return;
// 	// } else {
// 	// 	const data = window.URL.createObjectURL(blob);
// 	// 	let link = document.createElement('a');
// 	// 	link.href = data;
// 	// 	link.download = filename;
// 	// 	link.click();
// 	// 	setTimeout(() => {
// 	// 		window.URL.revokeObjectURL(data);
// 	// 	}, 100);
// 	// }
// 	let img = document.querySelector("#photo");
// 	img.src = window.URL.createObjectURL(data);
// };


function loadQuery() {
	const clientContext = new SP.ClientContext('/');
	const web = clientContext.get_web();
	const webs = web.get_webs();
	const lists = web.get_lists();
	const webCollection = clientContext.loadQuery(webs);
	const listsCollection = clientContext.loadQuery(lists);
	clientContext.executeQueryAsync(() => {
		const webs = clientContext.get_web().get_webs();
		const lists = clientContext.get_web().get_lists();
		const webCollection = clientContext.loadQuery(webs);
		const listsCollection = clientContext.loadQuery(lists);
		clientContext.executeQueryAsync(() => {
			console.log(webs);
			console.log(webCollection);
			console.log(lists);
			console.log(listsCollection);
		})
	})
}

// loadQuery()


// spx(['/a', '/b']).list(['c', 'd']).item('Title eq test').get()

const contexts = ['contextA', 'contextB', 'contextC'];

const lists = ['listA', 'listB', 'listC'];

const folders = ['a', 'b', 'c'];

const getSPObjectAsync = a => new Promise((resolve, reject) => setTimeout(() => resolve(`spObject${a}`), 100));

const mapContexts = fn => contexts => Promise.all(contexts.map(fn));

const mapLists = fn => lists => Promise.all(lists.map(fn));

const mapFolders = fn => folders => Promise.all(folders.map(fn));


const groupper = (by, array) => {
	const groupBys = typeOf(by) === 'array' ? by : [by];
	const groupFlat = by => array => {
		const groupped = {};
		array.map(el => {
			const elValue = el[by];
			const groupValue = groupped[elValue];
			groupped[elValue] = groupValue ? (typeOf(groupValue) === 'array' ? groupValue.concat(el) : [groupValue, el]) : el;
		})
		return groupped
	}

	const mapper = (array, fn) => {
		const result = {};
		const mapperR = (acc, el, prop) => {
			if (prop) {
				const childEl = el[prop];
				if (typeOf(childEl) === 'array') {
					acc[prop] = fn(childEl);
					return acc;
				} else {
					acc[prop] = childEl;
					if (typeOf(childEl) === 'object') {
						for (let childProp in childEl) {
							acc[prop] = childEl;
							mapperR(acc[prop], childEl, childProp);
						}
					}
					return acc;
				}
			} else {
				if (typeOf(el) === 'array') {
					return fn(el)
				} else {
					for (let prop in el) mapperR(acc, el, prop);
					return acc;
				}
			}
		}
		return mapperR(result, array);
	}

	const merger = (obj, fn) => groupBys.reduce((acc, el) => mapper(acc, fn(el)), Object.assign(obj));

	return merger(array, groupFlat);
}

// console.log(groupper(['a', 'b', 'c'], [{
// 	a: 1, b: 1, c: 1
// }, {
// 	a: 1, b: 1, c: 2
// }, {
// 	a: 1, b: 2, c: 1
// }, {
// 	a: 1, b: 2, c: 2
// }, {
// 	a: 2, b: 1, c: 1
// }, {
// 	a: 2, b: 1, c: 2
// }, {
// 	a: 2, b: 2, c: 1
// }, {
// 	a: 2, b: 2, c: 2
// }]));

window.deleteFolders = () => {
	const list = spx('/test/spx/testFolder16').list('test2');
	list.item().get().then(items => {
		console.time('item');
		list.item(items.map(item => item.ID)).delete({ noRecycle: true }).then((items) => {
			console.timeEnd('item');
			log(items);
		})
	})
}

window.createFolders = (serial) => {
	const foldersToCreate = [];
	for (let i = 0; i < 504; i++) {
		foldersToCreate.push({ Title: 'test' + i })
	}
	console.time('item');
	spx('/test/spx/testFolder16').list('test2').folder(foldersToCreate).create({ serial }).then((items) => {
		console.timeEnd('item');
		log(items);
	});
}

window.deleteWebFolders = () => {
	const web = spx('/test/spx/testFolder16');
	web.folder('Lists/test2/').get().then(folders => {
		console.log(folders);
		console.time('folders');
		web.folder(folders.map(folder => folder.ServerRelativeUrl)).delete({ noRecycle: true }).then((folders) => {
			console.timeEnd('folders');
			log(folders);
		})
	})
}

window.createWebFolders = (serial) => {
	const foldersToCreate = [];
	for (let i = 0; i < 1; i++) {
		foldersToCreate.push('Lists/test2/test' + i);
	}
	console.time('folders');
	spx('/test/spx/testFolder16').folder(foldersToCreate).create({ serial }).then((folders) => {
		console.timeEnd('folders');
		log(folders);
	});
}

window.createItems = (serial) => {
	const itemsToCreate = [];
	for (let i = 0; i < 504; i++) {
		itemsToCreate.push({ Title: 'test' + i })
	}
	console.time('item');
	spx('/test/spx/testFolder16').list('test2').item(itemsToCreate).create({ serial }).then((items) => {
		console.timeEnd('item');
		log(items);
	});
}

window.deleteItems = () => {
	const list = spx('/test/spx/testFolder16').list('test2');
	list.item().get().then(items => {
		console.time('item');
		list.item(items.map(item => item.ID)).delete({ noRecycle: true }).then((items) => {
			console.timeEnd('item');
			log(items);
		})
	})
}

window.update = function update() {
	const list = spx('/test/spx/testFolder16').list('test2');
	list.item().get().then(items => {
		console.time('item');
		list.item(items.map(item => {
			return item.ID;
			return {
				ID: item.ID,
				Title: 'l'
			}
		})).delete().then((items) => {
			console.timeEnd('item');
			log(items);
		})
	})
}

// update()

