import web from './test/web.js';
import folder from './test/folder.js';
import file from './test/file.js';
import list from './test/list.js';

let url = '/test/spx/testFolder';
let fileUrl = 'a/test.txt';
let destFolder = 'b';
let fileRelativeUrl = `${url}/${fileUrl}`;

export async function webBundle() {
	console.log('web start, total 6');
	console.log(1);
	await web.get().catch(async () => {});
	console.log(2);
	await web.create().catch(async () => {});
	console.log(3);
	await web.get();
	console.log(4);
	await web.update();
	console.log(5);
	await web.get();
	console.log(6);
	await web.delete();
	console.log('web success!');
}

export async function folderBundle() {
	console.log('folder start, total 4');
	console.log(1);
	await folder.get().catch(() => {})
	console.log(2);
	await folder.create();
	console.log(3);
	await folder.get();
	console.log(4);
	await folder.delete();
	console.log('folder success!');

}

export async function fileBundle() {
	console.log('file start, total 8');
	console.log(1);
	await file.create().catch(() => {});
	console.log(2);
	await file.get();
	console.log(3);
	await file.update();
	console.log(4);
	await file.getValue();
	console.log(5);
	await file.copy();
	console.log(6);
	await file.delete(destFolder);
	console.log(7);
	await file.move();
	console.log(8);
	await file.delete(destFolder);
	console.log('file success!');
}



// // web
// spx().web().get().then(log);
// spx('/').web('Alekseev').get().then(log);

// let alekseevWeb = spx('/Alekseev').web('Alekseev');

// alekseevWeb.get()

// // list
// spx('/Alekseev').list('test').get({cached:true}).then(log)

// // folder
// spx('/Alekseev').list('test').folder('a').get({cached:true}).then(log)
// spx('/Alekseev').list().folder('a').get({cached:true}).then(log)

// spx('/Alekseev').list('test').folder('a').get({cached:true}).then(log)
// spx('/Alekseev').folder('a').get({cached:true}).then(log)

// spx('/Alekseev/test').folder('a').get({cached:true}).then(log)
// spx('/Alekseev').folder('a').get({cached:true}).then(log)

// // item
// spx('/Alekseev').list('test').item('Title Eq test',{folder:'a'}).get({cached:true}).then(log)

// // file
// spx('/Alekseev').list('test').file('test.png',{folder:'a'}).get({cached:true}).then(log)
// spx('/Alekseev').list().file('test.png',{folder:'a'}).get({cached:true}).then(log)

// spx('/Alekseev/test').file('test.png',{folder:'a'}).get({cached:true}).then(log)
// spx('/Alekseev').file('test.png',{folder:'a'}).get({cached:true}).then(log)