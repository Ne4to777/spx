import spx from './../modules/site';

export default async () => {
  let site;
  site = await spx.get();
  console.log('get()');
  console.log(site);

  site = await spx.getCustomListTemplates();
  console.log('getCustomListTemplates()');
  console.log(site);

}