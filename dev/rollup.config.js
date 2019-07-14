import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
// import { terser } from 'rollup-plugin-terser'

const input = './src/modules/site.js'

const globals = {
	SP: 'SP',
	Microsoft: 'Microsoft',
	_spPageContextInfo: '_spPageContextInfo'
}

const external = ['axios', 'crypto-js']

const extensions = ['.js']

const plugins = [
	resolve({
		extensions
	}),
	babel({
		extensions,
		exclude: 'node_modules/**',
		runtimeHelpers: true
	})
]

const output = {
	globals,
	name: 'SPX',
	format: 'umd'
}

export default [
	{
		input,
		external,
		output: {
			...output,
			file: 'dist/bundle.js'
		},
		plugins
	},
	// {
	// 	input,
	// 	external,
	// 	output: {
	// 		...output,
	// 		file: 'dist/bundle.min.js'
	// 	},
	// 	plugins: [
	// 		...plugins,
	// 		terser()
	// 	]
	// }
]
