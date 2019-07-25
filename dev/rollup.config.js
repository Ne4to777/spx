import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
import { terser } from 'rollup-plugin-terser'

const input = './src/modules/web.js'

const globals = {
	SP: 'SP',
	Microsoft: 'Microsoft',
	window: 'window',
	axios: 'axios',
	'crypto-js': 'cryptoJS'
}

const external = ['axios', 'crypto-js']

const extensions = ['.js']

const plugins = [
	resolve({
		extensions,
		preferBuiltins: false
	}),
	babel({
		extensions,
		exclude: 'node_modules/**',
		runtimeHelpers: true
	}),
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
	{
		input,
		external,
		output: {
			...output,
			file: 'dist/bundle.min.js'
		},
		plugins: [
			...plugins,
			terser()
		]
	}
]
