const {
	deploy
} = require('aura-connector')
const path = require('path')
const config = require('./private.json')

deploy(path.win32.normalize(path.join(config.siteDisk, config.deployPath)))
