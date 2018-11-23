import * as cache from './../cache';
import { assert } from 'chai';
import { expect } from 'chai';
describe('Cache', () => {
  describe('getCache()', () => {
    const obj = cache.getCache();
    it('should be an object', () => {
      assert.isObject(obj);
    });
    it('should be empty', () => {
      expect(obj).to.be.empty
    });
  });
});