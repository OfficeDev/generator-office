const path = require('path');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');
const BrowserSyncPlugin = require('browser-sync-webpack-plugin');

module.exports = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',

    plugins: [
        new BrowserSyncPlugin(
            {
                https: true,
                host: 'localhost',
                port: 3000,
                proxy: 'https://localhost:3100/'
            },
            {
                reload: false
            }
        )
    ],

    devServer: {
        publicPath: '/',
        contentBase: path.resolve('dist'),
        https: true,
        compress: true,
        overlay: {
            warnings: false,
            errors: true
        },
        port: 3100,
        historyApiFallback: true
    }
});
