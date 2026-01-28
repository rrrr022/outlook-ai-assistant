const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');
const fs = require('fs');
const os = require('os');

const isProduction = process.env.NODE_ENV === 'production';
const DEFAULT_CLIENT_PORT = 3000;
const DEFAULT_SERVER_PORT = 3001;

// Get Office Add-in dev certs path
const certPath = path.join(os.homedir(), '.office-addin-dev-certs');
const httpsOptions = fs.existsSync(path.join(certPath, 'localhost.crt')) ? {
  key: fs.readFileSync(path.join(certPath, 'localhost.key')),
  cert: fs.readFileSync(path.join(certPath, 'localhost.crt')),
  ca: fs.readFileSync(path.join(certPath, 'ca.crt')),
} : true;

module.exports = {
  entry: {
    taskpane: './src/client/index.tsx',
  },
  output: {
    path: path.resolve(__dirname, 'dist/client'),
    filename: '[name].bundle.js',
    clean: true,
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    alias: {
      '@': path.resolve(__dirname, 'src'),
      '@client': path.resolve(__dirname, 'src/client'),
      '@server': path.resolve(__dirname, 'src/server'),
      '@shared': path.resolve(__dirname, 'src/shared'),
    },
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader'],
      },
      {
        test: /\.(png|jpg|jpeg|gif|svg|ico)$/,
        type: 'asset/resource',
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/client/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane'],
    }),
    new HtmlWebpackPlugin({
      template: './src/client/index.html',
      filename: 'index.html',
      chunks: [],
      inject: false,
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: 'src/client/assets',
          to: 'assets',
          noErrorOnMissing: true,
        },
        {
          from: 'manifest.azure.xml',
          to: 'manifest.azure.xml',
          noErrorOnMissing: true,
        },
        {
          from: 'staticwebapp.config.json',
          to: 'staticwebapp.config.json',
          noErrorOnMissing: true,
        },
        {
          from: 'public',
          to: '.',
          noErrorOnMissing: true,
        },
      ],
    }),
    // Define environment variables for client
    new webpack.DefinePlugin({
      'process.env.DEFAULT_SERVER_PORT': JSON.stringify(DEFAULT_SERVER_PORT),
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist/client'),
    },
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
    server: {
      type: 'https',
      options: httpsOptions,
    },
    port: DEFAULT_CLIENT_PORT,
    hot: true,
    open: false,
    // Redirect root to taskpane.html
    historyApiFallback: {
      rewrites: [
        { from: /^\/$/, to: '/taskpane.html' },
      ],
    },
    // Auto-find available port if default is in use
    onListening: function (devServer) {
      const port = devServer.server.address().port;
      console.log(`\n🎨 Frontend running on https://localhost:${port}/taskpane.html\n`);
    },
    // Try next port if current is in use
    client: {
      overlay: {
        errors: true,
        warnings: false,
      },
    },
  },
  devtool: isProduction ? 'source-map' : 'eval-source-map',
};
