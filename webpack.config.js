const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (_env, argv) => {
  const isDevelopment = argv.mode === "development";

  return {
    entry: {
      commands: "./src/commands/commands.ts",
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
        patterns: [{ from: "src/assets", to: "assets" }],
      }),
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
          headers: {
            "Access-Control-Allow-Origin": "*",
          },
        }
      : undefined,
  };
};
