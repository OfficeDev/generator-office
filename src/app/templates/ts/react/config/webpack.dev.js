const path = require('path');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

module.exports = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',

    devServer: {
        publicPath: '/',
        contentBase: path.resolve('dist'),
        https: true,
        compress: true,
        open: true,
        overlay: {
            warnings: false,
            errors: true
        },
        port: 3000,
        historyApiFallback: true
    }
});
