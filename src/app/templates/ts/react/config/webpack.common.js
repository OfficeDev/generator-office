const webpack = require('webpack');
const path = require('path');
const package = require('../package.json');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const autoprefixer = require('autoprefixer');

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
    vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react'
    ],
    app: [
        'react-hot-loader/patch',
        './index.tsx',
    ],
    'function-file': '../function-file/function-file.ts'
};

const rules = [
    {
        test: /\.tsx?$/,
        use: [
            'react-hot-loader/webpack',
            'ts-loader'
        ],
        exclude: /node_modules/
    },
    {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
    },
    {
        test: /\.less$/,
        use: ['style-loader', 'css-loader', 'less-loader']
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
    publicPath: '/',
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
            ],
            htmlLoader: {
                minimize: true
            }
        }
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
    optimization: {
        splitChunks: {
          chunks: 'async',
          minChunks: Infinity,
          name: 'vendor'
        }
      },
    plugins: [
        ...WEBPACK_PLUGINS,
        new ExtractTextPlugin('[name].[hash].css'),
        new HtmlWebpackPlugin({
            title: '<%= projectDisplayName %>',
            filename: 'index.html',
            template: './index.html',
            chunks: ['app', 'vendor', 'polyfills']
        }),
        new HtmlWebpackPlugin({
            title: '<%= projectDisplayName %>',
            filename: 'function-file/function-file.html',
            template: '../function-file/function-file.html',
            chunks: ['function-file']
        }),
        new CopyWebpackPlugin([
            {
                from: '../assets',
                ignore: ['*.scss'],
                to: 'assets',
            }
        ])
    ]
};