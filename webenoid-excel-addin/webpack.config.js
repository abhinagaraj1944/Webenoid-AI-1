/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: [
        "./src/taskpane/taskpane.js",
        "./src/taskpane/taskpane.html",
      ],
    },

    output: {
      clean: true,
    },

    resolve: {
      extensions: [".html", ".js"],
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/i,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },

    plugins: [
      // Taskpane HTML
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),

      // Static files
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              return dev
                ? content
                : content
                  .toString()
                  .replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
          {
            from: "./src/taskpane/assets",
            to: "assets",
          }
        ],
      }),
    ],

    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: 3000,
    },
  };
};
