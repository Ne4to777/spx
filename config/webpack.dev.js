const webpack = require('webpack');
const path = require('path');
const proxyPathConfig = {
	target: 'http://localhost:8080',
	secure: false
};
module.exports = {
	mode: 'development',
	devtool: 'inline-source-map',
	entry: ['./src/index.js'],
	output: {
		filename: 'index.js',
		path: path.resolve(__dirname, 'dist'),
		publicPath: '/'
	},
	devServer: {
		contentBase: './dist',
		port: 3000,
		hot: true,
		proxy: {
			'/_api': proxyPathConfig,
			'**/_api/**': proxyPathConfig,
			'/_vti_bin': proxyPathConfig,
			'**/_vti_bin/**': proxyPathConfig,
			'/_layouts/**': proxyPathConfig,
			'**/_layouts/**': proxyPathConfig
		}
	},
	plugins: [
		new webpack.HotModuleReplacementPlugin(),
		new webpack.ProvidePlugin({
			typeOf: [path.resolve(__dirname, '../lib/util.js'), 'typeOf'],
			log: [path.resolve(__dirname, '../lib/util.js'), 'log']
		})
	],
	performance: {
		hints: false
	},
	resolve: {
		alias: {
			spx: path.resolve(__dirname, './../dist/spx.js')
		}
	}
};