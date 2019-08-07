import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
import commonjs from 'rollup-plugin-commonjs'
import { terser } from 'rollup-plugin-terser'

const getBundle = (file, plugins = []) => ({
	input: './src/modules/web.js',
	output: [{
		globals: {
			axios: 'axios',
			'crypto-js': 'CryptoJS'
		},
		sourcemap: true,
		sourcemapExcludeSources: true,
		name: 'spx',
		file: `dist/${file}`,
		format: 'iife'
	}],
	external: ['axios', 'crypto-js'],
	plugins: [
		...plugins,
		resolve(),
		commonjs(),
		babel({
			exclude: 'node_modules/**',
			extensions: ['.js']
		})
	]
})

export default [
	// getBundle('index.js'),
	getBundle('index.min.js', [terser()])
]
