const RestProxy = require('sp-rest-proxy');

new RestProxy({
	configPath: './config/private.json',
	port: 8080
}).serve();