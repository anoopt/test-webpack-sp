var webpack = require('webpack');
var SPSaveWebpackPlugin = require('spsave-webpack-plugin');

module.exports = {
    
    entry: {
        vendor:[
                "react",
                "react-dom",
                "sp-pnp-js"
               ],
        app: ["./src/app.tsx"]
    },
    output: {
        filename: "test.webpack-sp.[name].js",
        path: __dirname + "/dist"
    },

    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: [".ts", ".tsx", ".js", ".json"]
    },

    module: {
        rules: [
            // All files with a '.ts' or '.tsx' extension will be handled by 'awesome-typescript-loader'.
            { 
                test: /\.tsx?$/, 
                loader: "awesome-typescript-loader",
                options: {
                    configFileName: "tsconfig.json"
                }
             }
        ]
    },
    plugins: [
        new webpack.optimize.CommonsChunkPlugin({
            name: "vendor",
            minChunks: Infinity
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
    ]
};