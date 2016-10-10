const path = require('path');
const webpack = require('webpack');

module.exports = {
    entry: "./webpack-sample/index.js",
    output: {
        path: path.join(__dirname, "webpack-sample"),
        filename: "bundle.js"
    },
    module: {
        loaders: [
            { test: /\.css$/, loader: "style!css" }
        ]
    },
    plugins: [
        new webpack.optimize.UglifyJsPlugin({
            compress: {
                warnings: false
            }
        })
    ],
    watch: false
};