const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const { readConfig } = require("./server/config");
const { createApp } = require("./server/app");

module.exports = async (_env, argv) => {
  const isDevelopment = argv.mode === "development";

  return {
    entry: {
      commands: "./src/commands/commands.ts",
      taskpane: "./src/taskpane/taskpane.ts",
      original: "./src/taskpane/original.ts",
      consent: "./src/taskpane/consent.ts",
    },
    devtool: isDevelopment ? "source-map" : false,
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        inject: "body",
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "src/assets", to: "assets" }, { from: "src/taskpane/taskpane.css", to: "taskpane.css" }],
      }),
      new HtmlWebpackPlugin({ filename: "taskpane.html", template: "./src/taskpane/taskpane.html", chunks: ["taskpane"], inject: "body" }),
      new HtmlWebpackPlugin({ filename: "original.html", template: "./src/taskpane/original.html", chunks: ["original"], inject: "body" }),
      new HtmlWebpackPlugin({ filename: "consent.html", template: "./src/taskpane/consent.html", chunks: ["consent"], inject: "body" }),
    ],
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },
    devServer: isDevelopment
      ? {
          server: {
            type: "https",
            options: await devCerts.getHttpsServerOptions(),
          },
          port: 3000,
          host: "localhost",
          hot: false,
          setupMiddlewares: (middlewares, devServer) => {
            devServer.app.use(createApp(readConfig()));
            return middlewares;
          },
        }
      : undefined,
  };
};
