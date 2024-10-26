import resolve from '@rollup/plugin-node-resolve'
import commonjs from '@rollup/plugin-commonjs'
import zip from 'rollup-plugin-zip'

import { chromeExtension, simpleReloader} from 'rollup-plugin-chrome-extension'

export default {
  input: 'manifest.json',
  output: {
    dir: 'build',
    format: 'esm',
  },
  cache: true, // required for zip plugin
  plugins: [
    // always put chromeExtension() before other plugins
    chromeExtension(),
    simpleReloader(),
    // the plugins below are optional
    resolve(),
    commonjs(),
    zip()
  ],
}