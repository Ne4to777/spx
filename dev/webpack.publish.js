const config = require('./private.json')
const path = require('path')

module.exports = {
  mode: 'production',
  devtool: 'source-map',
  entry: ['./src/modules/site.js'],
  output: {
    filename: config.filename,
    path: path.resolve(__dirname, '../publish'),
  }
}