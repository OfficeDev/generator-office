const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
    entry: {
        polyfill: 'babel-polyfill',
        app: './src/index.ts',
        'function-file': './function-file/function-file.ts'
    },
    
    resolve: {

        extensions: ['.ts', '.tsx', '.html', '.js']

    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                exclude: /node_modules/,
                use: 'babel-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif)$/,
                use: 'file-loader'
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['polyfill', 'app']
        }),
        new HtmlWebpackPlugin({
            template: './function-file/function-file.html',
            filename: 'function-file/function-file.html',
            chunks: ['function-file']
        })
    ]
};