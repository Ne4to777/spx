const getTermStore = clientContext => SP
	.Taxonomy
	.TaxonomySession
	.getTaxonomySession(clientContext)
	.getDefaultKeywordsTermStore()

const getTermSet = clientContext => getTermStore(clientContext).get_keywordsTermSet()


const getAllTerms = clientContext => getTermSet(clientContext).getAllTerms()

const getSPObject = (clientContext, elementUrl) => getAllTerms(clientContext).getByName(elementUrl)

const clientContext = new SP.ClientContext('http://localhost:3000')
const element = getTermStore(clientContext)

clientContext.load(element)
clientContext.executeQueryAsync(() => {
	console.log(element);
})
