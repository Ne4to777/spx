const config = require('./private.json')
const CopyPlugin = require('copy-webpack-plugin')

module.exports = {
	plugins: [
		new CopyPlugin(
			[
				{
					from: '**/*',
					to: config.path + '/project',
					ignore: ['node_modules/**']
				}
			],
			{
				copyUnmodified: true
			}
		)
	]
}
