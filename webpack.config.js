/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const glob = require("glob");
const urlDev = "https://localhost:3000/";
const urlProd = "https://www.excel.meekou.cn/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}
//const meekou = glob.sync("./src/shared/AppConsts.ts");
module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      // meekou: [
      //   "./src/shared/appconsts.ts",
      //   "./src/shared/meekouconsts.ts",
      //   "./src/shared/dialoginput.ts",
      //   "./src/services/meekouapi.ts",
      //   "./src/services/common.model.ts",
      //   "./src/services/http.ts",
      // ],
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      functions: "./src/functions/functions.ts",
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
      login: "./src/login/login.ts",
      dataFromWeb: "./src/pages/data-from-web/data-from-web.ts",
      dialog: "./src/dialog/dialog.ts",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"],
      }),
      new HtmlWebpackPlugin({
        filename: "login.html",
        template: "./src/login/login.html",
        chunks: ["polyfill", "login"],
      }),
      new HtmlWebpackPlugin({
        filename: "data-from-web.html",
        template: "./src/pages/data-from-web/data-from-web.html",
        chunks: ["polyfill", "dataFromWeb"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog/dialog.html",
        chunks: ["polyfill", "dialog"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/icon-*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      static: [__dirname],
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
