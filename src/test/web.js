import web from './../modules/web';
const test = _ => {
  web().get().then(res => log('root web', res));
  web('/').get().then(res => log('root webs', res));
  web('test/spx/functional1').get().then(res => log('single', res));
  web(['test/spx/functional1']).get().then(res => log('array of 1', res));
  web({ Url: 'test/spx/functional1' }).get().then(res => log('single', res));
  web([{ Url: 'test/spx/functional1' }]).get().then(res => log('array of 1', res));
  web(['test/spx/functional1', 'test/spx/functional2']).get().then(res => log('array of 2', res));
  web([{ Url: 'test/spx/functional1' }, { Url: 'test/spx/functional2' }]).get().then(res => log('array of 2', res));
}