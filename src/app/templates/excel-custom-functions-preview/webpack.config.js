const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
    entry: {
        customfunctions: ["./src/customfunctions.js"],
        customfunctionsjson: ["./config/customfunctions.json"],
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.html', '.js', 'json']
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                exclude: /node_modules/,
                use: 'ts-loader'
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
            chunks: ['customfunctions']
        }),
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['customfunctionsjson']
        })
    ],
    devServer: {
        port: 3000,
        hot: true,
        inline: true,
        headers: {
            "Access-Control-Allow-Origin": "*"
        }
    }
};