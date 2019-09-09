const robocopy = require('robocopy')

robocopy({
	source: './dev/assets',
	destination: './publish/assets/sp',
	files: ['sp.assembly.js']
})
robocopy({
	source: './dev/setupFiles',
	destination: './publish/assets/jest',
	files: ['sp.assembly.js']
})
