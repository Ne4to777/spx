const webpack = require('webpack')
const RestProxy = require('sp-rest-proxy')
const path = require('path')
const proxyPathConfig = {
	target: 'http://localhost:8080',
	secure: false
}
module.exports = {
	mode: 'development',
	devtool: 'inline-source-map',
	entry: [/* 'babel-polyfill',  */ './src/index.js'],
	// module: {
	// 	rules: [{
	// 		test: /\.js$/,
	// 		loader: 'babel-loader'
	// 	}]
	// },
	output: {
		filename: 'index.js',
		path: path.resolve(__dirname, 'dist'),
		publicPath: '/'
	},
	devServer: {
		contentBase: './public',
		port: 3000,
		hot: true,
		proxy: {
			'/_api': proxyPathConfig,
			'**/_api/**': proxyPathConfig,
			'/_vti_bin': proxyPathConfig,
			'**/_vti_bin/**': proxyPathConfig,
			'/_layouts/**': proxyPathConfig,
			'**/_layouts/**': proxyPathConfig
		},
		before: app =>
			new RestProxy({
				configPath: './dev/private.json',
				hostname: 'localhost',
				port: 8080
			}).serve()
	},
	plugins: [new webpack.HotModuleReplacementPlugin()],
	performance: {
		hints: false
	}
}
