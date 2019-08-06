import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
import commonjs from 'rollup-plugin-commonjs'
// import { terser } from 'rollup-plugin-terser'

// const plugins = [
// 	resolve({
// 		extensions,
// 		preferBuiltins: false
// 	}),
// 	babel({
// 		extensions,
// 		exclude: 'node_modules/**',
// 		// runtimeHelpers: true
// 	}),
// ]

export default {
	input: './src/modules/web.js',
	output: [{
		file: 'dist/bundle.cjs.js',
		format: 'cjs'
	},
	{
		file: 'dist/bundle.esm.js',
		format: 'esm'
	},
	{
		globals: {
			axios: 'axios',
			'crypto-js': 'CryptoJS'
		},
		name: 'spx',
		file: 'dist/bundle.umd.js',
		format: 'umd'
	},
	],
	external: ['axios', 'crypto-js'],
	plugins: [
		resolve(),
		commonjs(),
		babel({
			exclude: 'node_modules/**',
			extensions: ['.js']
		})
	]
}

// export default [
// 	{
// 		input,
// 		external,
// 		output: {
// 			...output,
// 			file: 'dist/bundle.js'
// 		},
// 		plugins
// 	},
// 	{
// 		input,
// 		external,
// 		output: {
// 			...output,
// 			file: 'dist/bundle.min.js'
// 		},
// 		plugins: [
// 			...plugins,
// 			terser()
// 		]
// 	}
// ]
