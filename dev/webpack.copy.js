const CopyPlugin = require('copy-webpack-plugin')
const config = require('./private.json')

module.exports = {
	mode: 'none',
	plugins: [
		new CopyPlugin(
			[
				{
					from: '../dist',
					to: config.path
				}
			],
			{
				copyUnmodified: true
			}
		)
	]
}
