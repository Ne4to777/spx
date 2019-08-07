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
		file: `publish/${file}`,
		format: 'umd'
	}],
	external: ['axios', 'crypto-js'],
	plugins: [
		...plugins,
		resolve(),
		commonjs(),
	]
})

export default [
	// getBundle('index.js'),
	getBundle('index.min.js', [terser()])
]
