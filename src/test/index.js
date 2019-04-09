import cache from '../test/cache';
import site from '../test/site';
import web from '../test/web';
import column from '../test/column';
import fileList from '../test/fileList';
import fileWeb from '../test/fileWeb';
import folderList from '../test/folderList';
import folderWeb from '../test/folderWeb';
import item from '../test/item';
import list from '../test/list';
import user from '../test/user';
import tag from '../test/tag';
import queryParser from '../test/query-parser';
import { testIsOk } from './../lib/utility'

export default async _ => {
  // await cache();
  // await site();
  // await web();
  // await folderWeb();
  // await fileWeb();
  // await list();
  // await column();
  // await folderList();
  // await fileList();
  // await item();
  // await tag();
  // await queryParser();
  testIsOk('whole test')()
}