const webpack = require('webpack');
const path = require('path');
const package = require('../package.json');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const { CheckerPlugin } = require('awesome-typescript-loader');
const autoprefixer = require('autoprefixer');
const perfectionist = require('perfectionist');

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
    vendor: './vendor.ts',
    app: [
        'webpack-dev-server/client?http://localhost:3000',
        'webpack/hot/only-dev-server',
        './main.tsx',
    ]
};

const rules = [
    {
        test: /\.tsx?$/,
        use: [
            'react-hot-loader',
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
        test: /\.scss$/,
        use: ExtractTextPlugin.extract({
            fallback: 'style-loader',
            use: ['css-loader', 'postcss-loader', 'sass-loader']
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
            postcss: [
                autoprefixer({ browsers: ['Safari >= 8', 'last 2 versions'] }),
                perfectionist
            ],
            htmlLoader: {
                minimize: true
            }
        }
    }),
    new webpack.optimize.CommonsChunkPlugin({
        name: ['vendor', 'app'],
        minChunks: 2
    })
];

module.exports = {
    context: path.resolve('./src'),
    entry,
    output,
    resolve: {
        extensions: ['.js', '.jsx', '.ts', '.tsx', '.scss', '.css', '.html']
    },
    module: {
        rules,
    },
    plugins: [
        ...WEBPACK_PLUGINS,
        new ExtractTextPlugin('[name].[hash].css'),
        new HtmlWebpackPlugin({
            title: '<%= projectDisplayName %>',
            filename: 'index.html',
            template: 'index.html',
            chunks: ['app', 'vendor', 'polyfills']
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