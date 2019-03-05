import web from './../modules/web';

const test = _ => {
  web.get().then(console.log)
  web.getWebTemplates().then(console.log)
  web.getCustomListTemplates().then(console.log)
}