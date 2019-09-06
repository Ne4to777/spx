import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
import commonjs from 'rollup-plugin-commonjs'
import { terser } from 'rollup-plugin-terser'

const getBundle = (file, format, plugins = []) => ({
	input: './src/modules/web.js',
	output: [{
		globals: {
			axios: 'axios',
			'crypto-js': 'CryptoJS'
		},
		sourcemap: true,
		sourcemapExcludeSources: true,
		name: 'spx',
		file,
		format
	}],
	external: ['axios', 'crypto-js'],
	plugins: [
		...plugins,
		resolve(),
		commonjs()
	]
})

const babelPlugin = babel({
	exclude: 'node_modules/**',
	extensions: ['.js']
})
const terserPlugin = terser()
export default [
	getBundle('publish/iife/index.js', 'iife', [babelPlugin]),
	getBundle('publish/iife/index.min.js', 'iife', [babelPlugin, terserPlugin]),
	getBundle('publish/umd/index.js', 'umd'),
	getBundle('publish/umd/index.min.js', 'umd', [terserPlugin])
]
