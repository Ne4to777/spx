import {
  AbstractBox,
  prepareResponseJSOM,
  load,
  executorJSOM,
  methodEmpty,
  getInstance,
  switchCase,
  typeOf,
  ifThen,
  isArrayFilled,
  map,
  constant,
  deep4Iterator,
} from './../lib/utility';


// Internal

const NAME = 'attachment';

const getSPObject = element => itemSPObject => {
  const url = element.Url;
  const attachments = getSPObjectCollection(itemSPObject);
  return url ? attachments.getByFileName(url) : attachments;
};
const getSPObjectCollection = methodEmpty('get_attachmentFiles');

const liftItemType = switchCase(typeOf)({
  object: item => Object.assign({}, item),
  string: (item = '') => ({ Url: item }),
  default: _ => ({ Url: void 0 })
})

class Box extends AbstractBox {
  constructor(value) {
    super(value);
    this.value = this.isArray
      ? ifThen(isArrayFilled)([
        map(liftItemType),
        constant([liftItemType()])
      ])(value)
      : liftItemType(value);
  }
}

// Inteface

export default parent => elements => {
  const instance = {
    box: getInstance(Box)(elements),
    NAME,
    parent
  }

  const iterator = deep4Iterator({
    contextBox: instance.parent.parent.parent.box,
    listBox: instance.parent.parent.box,
    parentBox: instance.parent.box,
    elementBox: instance.box
  });


  return {
    get: async (opts = {}) => {
      const { clientContexts, result } = await iterator(({ clientContext, listElement, parentElement, element }) => {
        const contextSPObject = instance.parent.parent.parent.getSPObject(clientContext);
        const listSPObject = instance.parent.parent.getSPObject(listElement.Url)(contextSPObject);
        const itemSPObject = instance.parent.getSPObject(parentElement)(listElement.Url)(listSPObject);
        const spObject = getSPObject(element)(itemSPObject);
        return load(clientContext)(spObject)(opts)
      })
      await instance.parent.parent.parent.box.chain(el => Promise.all(clientContexts[el.Url].map(clientContext => executorJSOM(clientContext)(opts))))
      return prepareResponseJSOM(opts)(result)
    }
  }
}

