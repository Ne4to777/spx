const webpack = require('webpack')
const RestProxy = require('sp-rest-proxy')
const path = require('path')

module.exports = {
	mode: 'development',
	devtool: 'inline-source-map',
	entry: ['./src/index.js'],
	output: {
		filename: 'index.js',
		path: path.resolve(__dirname, 'dist'),
	},
	devServer: {
		contentBase: './public',
		port: 3000,
		hot: true,
		before: (app) => {
			new RestProxy({
				configPath: './dev/private.json',
				hostname: 'localhost',
				port: 8080
			}, app).serveProxy()
		}
	},
	plugins: [new webpack.HotModuleReplacementPlugin()],
	performance: {
		hints: false
	}
}
