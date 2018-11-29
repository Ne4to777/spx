const path = require('path');
const webpack = require('webpack');

module.exports = {
	mode: 'production',
	devtool: 'source-map',
	entry: ['babel-polyfill', './src/modules/site.js'],
	module: {
		rules: [{
			test: /\.js$/,
			loader: 'babel-loader'
		}]
	},
	plugins: [
		new webpack.ProvidePlugin({
			typeOf: [path.resolve(__dirname, '../lib/util.js'), 'typeOf'],
			log: [path.resolve(__dirname, '../lib/util.js'), 'log']
		})
	],
	output: {
		filename: 'spx.js',
		path: path.resolve(__dirname, '../dist'),
		library: 'SPX',
		libraryTarget: 'umd'
	}
};