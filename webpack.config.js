const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");

module.exports = async (env) => ({
  entry: {},
  output: {
    path: path.resolve(__dirname, "dist"),
    clean: true,
  },
  plugins: [
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: "src", to: "src" },
        { from: "manifest*.{json,xml}", to: "[name][ext]" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.resolve(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: env.WEBPACK_BUILD ? undefined : await devCerts.getHttpsServerOptions(),
    },
    port: process.env.npm_package_config_dev_server_port || 3000,
  },
});
