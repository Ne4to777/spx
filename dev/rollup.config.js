import babel from 'rollup-plugin-babel'
import resolve from 'rollup-plugin-node-resolve'
import commonjs from 'rollup-plugin-commonjs'
import { terser } from 'rollup-plugin-terser'

const isBuild = process.env.packingType === 'build'

const getBundle = (file, format, plugins = []) => ({
	input: './src/modules/web.js',
	output: [{
		globals: {
			axios: 'axios',
			'crypto-js': 'CryptoJS'
		},
		sourcemap: true,
		name: 'spx',
		file,
		format
	}],
	external: ['axios', 'crypto-js'],
	plugins: [
		...plugins,
		resolve({
			mainFields: ['browser']
		}),
		commonjs()
	]
})

const babelPlugin = babel({
	exclude: 'node_modules/**',
	extensions: ['.js']
})
const terserPlugin = terser()

const buildBundle = () => getBundle('dist/index.min.js', 'iife', [babelPlugin, terserPlugin])

const publishBundle = () => [
	getBundle('publish/iife/index.js', 'iife', [babelPlugin]),
	getBundle('publish/iife/index.min.js', 'iife', [babelPlugin, terserPlugin]),
	getBundle('publish/umd/index.js', 'umd'),
	getBundle('publish/umd/index.min.js', 'umd', [terserPlugin])
]
export default isBuild ? buildBundle() : publishBundle()
