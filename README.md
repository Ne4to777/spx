# spx

## Installation
```sh
$ npm i
```

## Serve
```sh
$ npm run serve
```

Enter authorization data on first launch

## Build
```sh
$ npm run build
```

## Deploy
```sh
$ npm run deploy
```

## API

### Options
Options for async calls
```javascript
const opts = {
	view: ['Title', 'ID'], // (Array | string) show only 'Title' and 'ID' fields in response
	groupBy: ['Title', 'Author'], // (Array | string) response grouped first by 'Title' and then by 'Author' field. Returns object of arrays or objects.
	mapBy: 'Title', // (string) same as groupBy, but duplicated keys are merged. Always returns object of objects
	expanded: true, // (boolean) response as SP constructed object or collection (like SP.Item)
	noRecycle: true, // (boolean) delete element bypassing recycle bin. Only for .delete() operation.
	silent: true, // (boolean) all console messages are turned off
	silentErrors: true, // (boolean) error console messages are turned off
	silentInfo: true, // (boolean) info console messages are turned off
	detailed: true, // (boolean) detailed console report on create/update/delete operations
	asItem: true, // (boolean) get files and folders as items. Only for .folder() and .file()
	showCaml: true, // (boolean) show XML parsed CAML query in console. Only for .item()
}
```

### Web
```javascript
// get: object => Promise<object|Array>
spx().get(opts) // get root web
spx('foo').get(opts) // get subweb with Url 'foo'
spx('foo/bar').get(opts) // get subweb with Url 'bar' nested in 'foo'
spx('/').get(opts) // get root webs collection
spx('foo/').get(opts) // get subwebs collection in 'foo'

// create: object => Promise<object|Array>
spx({Url:'foo'}).create(opts) // creates subweb with Url and Title 'foo'. You can specify any editable field
spx('foo').create(opts) // same as previous
spx({Url:'foo/bar'}).create(opts) // creates subweb with Url 'foo/bar' and Title 'bar'
spx('foo/bar').create(opts) // same as previous
spx([{Url:'foo'},{Url:'bar'}]).create(opts) // creates 'foo' and 'bar' subwebs in root web
spx(['foo','bar']).create(opts) // same as previous

// update: object => Promise<object|Array>
spx({Url:'foo', Title: 'new foo'}).update(opts) // updates subweb Title with Url 'foo'. You can update any editable field
spx({Url:'foo/bar', Title: 'new bar'}).update(opts) // updates subweb Title with Url 'foo/bar'
spx([{Url:'foo', Title: 'new foo'},{Url:'bar', Title: 'new bar'}]).update(opts) // updates 'foo' and 'bar' subwebs in root web

// delete: object => Promise<object|Array>
spx({Url:'foo'}).delete(opts) // deletes subweb with Url 'foo'.
spx('foo').delete(opts) // same as previous
spx({Url:'foo/bar'}).delete(opts) // deletes subweb with Url 'foo/bar'
spx('foo/bar').delete(opts) // same as previous
spx([{Url:'foo'},{Url:'bar'}]).delete(opts) // deletes 'foo' and 'bar' subwebs in root web
spx(['foo','bar']).delete(opts) // same as previous

// doesUserHavePermissions: string => boolean
spx('foo').doesUserHavePermissions('viewListItems') // checks user for viewListItems permissions available on foo web. Mask types: https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee556747(v=office.14)

// getPermissions: void => object
spx('foo').getPermissions() // get all user permissions for foo web

// breakRoleInheritance: void => object
spx('foo').breakRoleInheritance() // breakes role inheritance for foo web

// resetRoleInheritance: void => object
spx('foo').resetRoleInheritance() // restore role inheritance for foo web

// getSite: object => object
spx('foo').getSite(opts) // get site for foo web

// getCustomListTemplates: object => array
spx('foo').getCustomListTemplates(opts) // get custom list templates for foo web

// getWebTemplates: object => array
spx('foo').getWebTemplates(opts) // get web templates for foo web
```