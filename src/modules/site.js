import Web from './../modules/web';
import RecycleBin from './../modules/recycleBin';
import * as cache from './../cache';
import * as utility from './../utility';
import User from './../modules/user';
import Tag from './../modules/tag';

const site = contextUrl => new Web(contextUrl);

// Interface

site.user = elementUrl => new User(this, elementUrl);
site.tag = elementUrl => new Tag(this, elementUrl);

site.get = async opts => execute(spObject => (spObject.cachePath = 'properties', spObject), opts);

site.getCustomListTemplates = async opts =>
  execute(spContextObject => {
    const spObject = spContextObject.getCustomListTemplates(spContextObject.get_context().get_web());
    spObject.cachePath = 'customListTemplates';
    return spObject
  }, opts);

// Internal

site._name = 'site';

site._contextUrls = ['/'];

const execute = async (spObjectGetter, opts = {}) => {
  const { cached } = opts;
  const clientContext = utility.getClientContext('/');
  const spObject = spObjectGetter(clientContext.get_site());
  const cachePaths = ['site', spObject.cachePath];
  const spObjectCached = cached ? cache.get(cachePaths) : null;
  if (cached && spObjectCached) {
    return utility.prepareResponseJSOM(spObjectCached, opts)
  } else {
    const spObjects = utility.load(clientContext, spObject, opts);
    await utility.executeQueryAsync(clientContext, opts);
    cache.set(cachePaths, spObjects);
    return utility.prepareResponseJSOM(spObjects, opts)
  }
}

site._getSPObject = clientContext => clientContext.get_site();

site.recycleBin = new RecycleBin(site);
window.spx = site;
export default site

site.createWebProject = async webPath => {
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