const config = require('./private.json')

module.exports = {
	mode: 'production',
	devtool: 'source-map',
	entry: ['babel-polyfill', './src/modules/web.js'],
	module: {
		rules: [
			{
				test: /\.js$/,
				loader: 'babel-loader'
			}
		]
	},
	output: {
		filename: config.filename,
		path: config.path,
		library: config.library,
		libraryTarget: 'umd'
	}
}
