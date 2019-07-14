const path = require('path')
const config = require('./private.json')

module.exports = {
	mode: 'production',
	devtool: 'source-map',
	entry: ['babel-polyfill', './src/modules/site.js'],
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
		path: path.resolve(__dirname, '../build'),
		library: config.library,
		libraryTarget: 'umd'
	}
}
