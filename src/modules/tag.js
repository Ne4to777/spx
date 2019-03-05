import {
  Box,
  load,
  getClientContext,
  getInstance,
  prepareResponseJSOM,
  executorJSOM,
  isString
} from './../utility'

// Interface

// const sessionTermSPObject = sessionSPObject.getTerm('c80b4506-8930-47c5-962b-4f371f9d9698');
// const termSPObject = termsSPObject.getByName('Авиация');

export default elements => {
  const instance = {
    box: getInstance(Box)(elements)
  }
  return {
    get: async opts => {
      const clientContext = getClientContext('/');
      const sessionSPObject = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext);
      const keywordsStoreSPObject = sessionSPObject.getDefaultKeywordsTermStore();
      const termSetSPObject = keywordsStoreSPObject.get_keywordsTermSet();
      const termsSPObject = termSetSPObject.getAllTerms();
      const elements = instance.box.chain(element => load(clientContext)(termsSPObject.getByName(element.Url))(opts))
      await executorJSOM(clientContext)(opts);
      return prepareResponseJSOM(opts)(elements);
    },
    search: (instance => async (opts = {}) => {
      const clientContext = getClientContext('/');
      const elements = instance.box.chain(element => {
        const lmi = SP.Taxonomy.LabelMatchInformation.newObject(clientContext);
        lmi.set_termLabel(element.Url);
        lmi.set_defaultLabelOnly(true);
        lmi.set_trimUnavailable(true);
        lmi.set_stringMatchOption(SP.Taxonomy.StringMatchOption[opts.exact ? 'exactMatch' : 'startsWith']);
        lmi.set_resultCollectionSize(opts.limit || 10);

        const spObject = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext).getDefaultKeywordsTermStore().get_keywordsTermSet().getTerms(lmi);
        return load(clientContext)(spObject)(opts);
      })
      await executorJSOM(clientContext)(opts);
      return prepareResponseJSOM(opts)(elements)
    })(instance),
    update: opts => {
      var keyword;
      var keywordsMerged = [];
      const elements = instance.box.chain(element => {

      })
      isString(listData) && (listData = this.getListData(listData));
      var clientContext = new SP.ClientContext(listData.path);
      for (var i = 0; i < keywords.length; i++) {
        keyword = keywords[i];
        keywordsMerged.push('-1;#' + keyword.value + '|' + keyword.id);
      }
      var list = clientContext.get_web().get_lists().getByTitle(listData.title);
      var listItem = list.getItemById(id);
      var field = list.get_fields().getByInternalNameOrTitle(columnName);
      var txField = clientContext.castTo(field, SP.Taxonomy.TaxonomyField);
      var termValues = new SP.Taxonomy.TaxonomyFieldValueCollection(clientContext, keywordsMerged.join(';#'), txField);
      txField.setFieldValueByValueCollection(listItem, termValues);
      listItem.update();
      clientContext.load(listItem);
      this.execute(clientContext, function () {
        console.log('updated keywords: ' + id);
        cb && cb(listItem);
      }, this.onQueryFailed);
    }
  }
}