const clientContext = new SP.ClientContext('http://localhost:3000')
const site = clientContext.get_site()

clientContext.load(site)
clientContext.executeQueryAsync(() => {
	console.log(site);
})