const config = require('./private.json')
const path = require('path')

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
		path: path.resolve(__dirname, '../publish'),
		library: config.library,
		libraryTarget: 'umd'
	}
}
