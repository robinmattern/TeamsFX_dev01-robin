// Copyright (c) Wictor Wilén. All rights reserved. 
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const webpack = require("webpack");
const nodeExternals = require("webpack-node-externals");

const ForkTsCheckerWebpackPlugin = require("fork-ts-checker-webpack-plugin");

const path = require("path");
const fs = require("fs");
const argv = require("yargs").argv;

const debug = argv.debug !== undefined;
const lint = !(argv["no-linting"] || argv.l === true);

const config = [{
    entry: {
        server: [
            path.join(__dirname, "/src/server/server.ts")
        ]
    },
    mode: debug ? "development" : "production",
    output: {
        path: path.join(__dirname, "/dist"),
        filename: "[name].js",
        devtoolModuleFilenameTemplate: debug ? "[absolute-resource-path]" : "[]"
    },
    externals: [nodeExternals()],
    devtool: debug ? "source-map" : "source-map",
    resolve: {
        extensions: [".ts", ".tsx", ".js"],
        alias: {}
    },
    target: "node",
    node: {
        __dirname: false,
        __filename: false
    },
    module: {
        rules: [{
            test: /\.tsx?$/,
            exclude: /node_modules/,
            use: {
                loader: "ts-loader",
                options: {
                    transpileOnly: true
                }
            }
        }]
    },
    plugins: [
        new ForkTsCheckerWebpackPlugin({
            typescript: {
                configFile: "./src/server/tsconfig.json"
            }
        })
    ]
},
{
    entry: {
        client: [
            path.join(__dirname, "/src/client/client.ts")
        ]
    },
    mode: debug ? "development" : "production",
    output: {
        path: path.join(__dirname, "/dist/web/scripts"),
        filename: "[name].js",
        libraryTarget: "umd",
        library: "tabMeetingStageviewVoteMovieFluid",
        publicPath: "/scripts/"
    },
    externals: {},
    devtool: debug ? "source-map" : "source-map",
    resolve: {
        extensions: [".ts", ".tsx", ".js"],
        alias: {},
        fallback: { "assert": false }
    },
    target: "web",
    module: {
        rules: [{
            test: /\.tsx?$/,
            exclude: /node_modules/,
            use: {
                loader: "ts-loader",
                options: {
                    transpileOnly: true
                }
            }
        }]
    },
    plugins: [
        new webpack.EnvironmentPlugin({ PUBLIC_HOSTNAME: undefined, 
            TAB_APP_ID: null, 
            TAB_APP_URI: null, 
            REACT_APP_TENANT_ID: null,
            REACT_APP_ORDERER: null,
            REACT_APP_STORAGE: null,
            REACT_APP_AZURETOKENURL: null,
            REACT_APP_REDIRECTURI: null }),
        new ForkTsCheckerWebpackPlugin({
            typescript: {
                configFile: "./src/client/tsconfig.json"
            }
        })
    ],
    devServer: {
        hot: false,
        host: "localhost",
        port: 9000,
        allowedHosts: "all",
        client: {
            overlay: {
                warnings: false,
                errors: true
            }
        },
        devMiddleware: {
            writeToDisk: true,
            stats: {
                all: false,
                colors: true,
                errors: true,
                warnings: true,
                timings: true,
                entrypoints: true
            }
        }
    }
}
];

module.exports = config;
