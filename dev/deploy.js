const robocopy = require('robocopy')
const config = require('./private.json')

robocopy({
	source: './publish/iife',
	destination: config.commitPath,
	files: ['index.min.js', 'index.min.js.map']
})
