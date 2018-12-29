import * as utility from './../utility';

export default class Tag {
  constructor(parent, elementUrl) {
    this._parent = parent;
    this._elementUrl = elementUrl;
    this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
  }
  async get() {
    const clientContext = utility.getClientContext('/');
    const session = SP.Taxonomy.TaxonomySession.getTaxonomySession(clientContext);
    const store = session.getDefaultKeywordsTermStore();
    const termSet = store.get_keywordsTermSet();
    const terms = termSet.getAllTerms();
    const term = terms.getByName('CMS');

    utility.load(clientContext, session);
    utility.load(clientContext, store);
    utility.load(clientContext, termSet);
    utility.load(clientContext, terms);
    utility.load(clientContext, term);
    await utility.executeQueryAsync(clientContext);
    console.log(session);
    console.log(utility.prepareResponseJSOM(session));
    console.log(store);
    console.log(utility.prepareResponseJSOM(store));
    console.log(termSet);
    console.log(utility.prepareResponseJSOM(termSet));
    console.log(terms);
    console.log(utility.prepareResponseJSOM(terms));
    console.log(term);
    console.log(utility.prepareResponseJSOM(term));

  }
}