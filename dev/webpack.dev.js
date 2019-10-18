const webpack = require('webpack')
const HtmlWebpackPlugin = require('html-webpack-plugin')
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
		contentBase: './dev/assets',
		port: 3000,
		hot: true,
		watchContentBase: true,
		before: (app) => {
			new RestProxy({
				configPath: './dev/private.json',
				hostname: 'localhost',
				port: 8080
			}, app).serveProxy()
		}
	},
	plugins: [
		new webpack.HotModuleReplacementPlugin(),
		new HtmlWebpackPlugin({
			template: 'dev/assets/index.ejs',
			templateParameters: {
				sp: 'sp.assembly.js',
				style: 'index.css'
			}
		})
	],
	performance: {
		hints: false
	}
}
