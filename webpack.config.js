const path = require('path');
const webpack = require('webpack');

module.exports = {
    entry: './lib/SharePoint.js',
    output: {
        filename: 'index.js'
    },
    node: {
        fs: 'empty'
    },
    module: {
        rules: [
            {
                test: /\.(js|jsx)$/,
                use: ['babel-loader']
            }
        ]
    }
};
