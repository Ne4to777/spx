const RestProxy = require('sp-rest-proxy');

new RestProxy({
	configPath: './dev/private.json',
	port: 8080
}).serve();