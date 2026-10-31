const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const pkg = require("./package.json");

// URL gốc của add-in: dev chạy local, production host trên GitHub Pages (BUSINESS_SPEC Q5).
const URL_DEV = "https://localhost:3000/";
const URL_PROD = "https://khoacrab-97.github.io/Addin_excel_NKTC/";

async function getHttpsOptions() {
  const devCerts = require("office-addin-dev-certs");
  const options = await devCerts.getHttpsServerOptions();
  return { ca: options.ca, key: options.key, cert: options.cert };
}

module.exports = async (env, argv) => {
  const isDev = argv.mode !== "production";
  // Chỉ cần chứng chỉ HTTPS khi chạy dev server (webpack serve), không cần khi build.
  const isServe = Boolean(env.WEBPACK_SERVE);

  return {
    devtool: isDev ? "source-map" : false,
    // Excel 2016/2019 chạy add-in trên Internet Explorer 11 webview → output ES5.
    target: ["web", "es5"],
    entry: {
      taskpane: "./src/taskpane/index.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      // Production: tên file có mã băm nội dung → Excel không dùng bản cũ trong cache sau khi cập nhật.
      filename: isDev ? "[name].js" : "[name].[contenthash:8].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.(png|jpe?g|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext]" },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        __APP_VERSION__: JSON.stringify(pkg.version),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*.png", to: "assets/[name][ext]" },
          // Biểu mẫu và file mẫu tải về từ tab Tiện ích (P3.1).
          { from: "templates/*.xlsx", to: "templates/[name][ext]" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              // Phiên bản manifest luôn khớp package.json (x.y.z → x.y.z.0).
              const versioned = content
                .toString()
                .replace(/<Version>[^<]*<\/Version>/, `<Version>${pkg.version}.0</Version>`);
              return isDev ? versioned : versioned.split(URL_DEV).join(URL_PROD);
            },
          },
          {
            // Thông tin phiên bản đang phát hành (dùng cho kiểm tra cập nhật — Phase 3).
            from: "package.json",
            to: "version.json",
            transform() {
              return JSON.stringify({ version: pkg.version, builtAt: new Date().toISOString() }, null, 2);
            },
          },
          {
            from: "src/site/index.html",
            to: "index.html",
            transform(content) {
              return content.toString().split("__APP_VERSION__").join(pkg.version);
            },
          },
          // GitHub Pages: tắt Jekyll để phục vụ nguyên trạng mọi file.
          { from: "src/site/.nojekyll", to: ".nojekyll", toType: "file" },
        ],
      }),
    ],
    performance: {
      // Fluent UI v8 làm bundle lớn hơn ngưỡng mặc định 244 KiB.
      maxEntrypointSize: 2 * 1024 * 1024,
      maxAssetSize: 2 * 1024 * 1024,
    },
    devServer: {
      port: Number(process.env.npm_package_config_dev_server_port || 3000),
      headers: { "Access-Control-Allow-Origin": "*" },
      server: isServe ? { type: "https", options: await getHttpsOptions() } : undefined,
    },
  };
};
