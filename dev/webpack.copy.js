const CopyPlugin = require('copy-webpack-plugin')
const config = require('./private.json')

module.exports = {
	plugins: [
		new CopyPlugin(
			[
				{
					from: '**/*',
					to: `${config.path}/project`,
					ignore: ['node_modules/**']
				}
			],
			{
				copyUnmodified: true
			}
		)
	]
}
