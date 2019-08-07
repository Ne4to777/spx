const robocopy = require('robocopy')

robocopy({
	source: './dist',
	destination: './publish',
})
