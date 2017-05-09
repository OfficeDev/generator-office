const webpack = require('webpack');
const path = require('path');
const package = require('../package.json');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const { CheckerPlugin } = require('awesome-typescript-loader');

const build = (() => {
    const timestamp = new Date().getTime();
    return {
        name: package.name,
        version: package.version,
        timestamp: timestamp,
        author: package.author
    };
})();

const entry = {
    app: './app.ts'
};

const rules = [
    {
        test: /\.ts$/,
        use: [
            'awesome-typescript-loader'
        ],
        exclude: /node_modules/
    },
    {
        test: /\.css$/,
        use: ExtractTextPlugin.extract({
            fallback: 'style-loader',
            use: 'css-loader'
        })
    },
    {
        test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
        use: {
            loader: 'file-loader',
            query: {
                name: 'assets/[name].[ext]'
            }
        }
    }
];

const output = {
    path: path.resolve('dist'),
    filename: '[name].[hash].js',
    chunkFilename: '[id].[hash].chunk.js'
};

const WEBPACK_PLUGINS = [
    new webpack.NamedModulesPlugin(),
    new webpack.NoEmitOnErrorsPlugin(),
    new webpack.BannerPlugin({ banner: `${build.name} v.${build.version} (${build.timestamp}) © ${build.author}` }),
    new webpack.DefinePlugin({
        ENVIRONMENT: JSON.stringify({
            build: build
        })
    }),
    new webpack.LoaderOptionsPlugin({
        options: {
            htmlLoader: {
                minimize: true
            }
        }
    }),
    new webpack.optimize.CommonsChunkPlugin({
        name: ['app'],
        minChunks: 1
    })
];

module.exports = {
    context: path.resolve('./src'),
    entry,
    output,
    resolve: {
        extensions: ['.js', '.ts', '.css', '.html']
    },
    module: {
        rules,
    },
    plugins: [
        ...WEBPACK_PLUGINS,
        new ExtractTextPlugin('[name].[hash].css'),
        new HtmlWebpackPlugin({
            title: '<%= name %>',
            filename: build.name,
            template: 'index.html',
            chunks: ['app']
        }),
        new CheckerPlugin(),
        new CopyWebpackPlugin([
            {
                from: './assets',
                ignore: ['*.scss'],
                to: 'assets',
            }
        ])
    ]
};