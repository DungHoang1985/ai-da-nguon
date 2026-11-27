/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { CspPlugin } = require("./csp-plugin");
const { DefinePlugin } = require("webpack");
const pkg = require("./package.json");

const urlDev = "https://localhost:3000/";
const urlProd = "https://dunghoang1985.github.io/ai-da-nguon/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
      // Trang mở trong hộp thoại của Office khi bấm "Đăng nhập" (Đợt 2)
      "dang-nhap": "./src/dialog/dang-nhap.ts",
    },
    output: {
      clean: true,
      filename: dev ? "[name].js" : "[name].[contenthash:8].js",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      // pptxgenjs chỉ dùng các mô-đun Node này khi chạy ngoài trình duyệt
      fallback: { fs: false, https: false, http: false, path: false, os: false, zlib: false, stream: false, url: false, crypto: false },
    },
    module: {
      rules: [
        // B16 (rà soát kiến trúc 20/11/2026, làm 26/11): chỉ ts-loader. Trước đây tệp .ts còn qua thêm babel-loader (.tsx thì không):
        // một lượt dịch thừa — đo 26/11: bỏ đi gói taskpane không đổi kích thước (Babel 8 mặc định đích trình duyệt hiện đại).
        // Đích dịch theo tsconfig.json.
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // Số phiên bản duy nhất: lấy từ package.json (src/core/version.ts)
      new CspPlugin({ dev }),
      new DefinePlugin({ __APP_VERSION__: JSON.stringify(pkg.version) }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane", "react"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "node_modules/pdfjs-dist/build/pdf.worker.min.mjs",
            to: "pdf.worker.min.js",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
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
      new HtmlWebpackPlugin({
        filename: "dang-nhap.html",
        template: "./src/dialog/dang-nhap.html",
        chunks: ["dang-nhap"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
