const Merge = require('webpack-merge');
const CommonConfig = require('./webpack.common.js');
var Visualizer = require('webpack-visualizer-plugin');
var SPSaveWebpackPlugin = require('spsave-webpack-plugin');

module.exports = Merge(CommonConfig, {
    devtool: "source-map",
    watch: true,
    plugins: [
        new Visualizer({
            filename: './statistics.dev.html'
        }),
        new SPSaveWebpackPlugin({
            "coreOptions": {
                "checkin": true,
                "checkinType": 1,
                "siteUrl": ""
            },
            "credentialOptions": {
                username: '',
                password: ''
            },
            "fileOptions": {
                "folder": "Style Library/test-webpack-sharepoint"
            }
        })
    ],
    module: {
        rules: [
            // All output '.js' files will have any sourcemaps re-processed by 'source-map-loader'.
            {
                enforce: "pre",
                test: /\.js$/,
                loader: "source-map-loader"
            }
        ]
    }
});