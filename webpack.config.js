const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, options) => {
    const isProd = options && options.mode === 'production';

    return {
        entry: {
            taskpane: './src/taskpane/taskpane.js',
            commands: './src/commands/commands.js',
        },
        output: {
            path: path.resolve(__dirname, 'dist'),
            filename: '[name].bundle.js',
            clean: true,
        },
        module: {
            rules: [
                {
                    test: /\.js$/,
                    exclude: /node_modules/,
                    use: {
                        loader: 'babel-loader',
                        options: {
                            presets: [['@babel/preset-env', {
                                targets: { browsers: ['last 2 versions', 'not IE 11'] },
                            }]],
                        },
                    },
                },
                {
                    test: /\.css$/,
                    use: [MiniCssExtractPlugin.loader, 'css-loader'],
                },
            ],
        },
        plugins: [
            new HtmlWebpackPlugin({
                filename: 'taskpane.html',
                template: './src/taskpane/taskpane.html',
                chunks: ['taskpane'],
                inject: 'body',
            }),
            new HtmlWebpackPlugin({
                filename: 'commands.html',
                template: './src/commands/commands.html',
                chunks: ['commands'],
                inject: 'body',
            }),
            new MiniCssExtractPlugin({ filename: '[name].css' }),
            new CopyWebpackPlugin({
                patterns: [
                    { from: 'assets', to: 'assets', noErrorOnMissing: true },
                    { from: 'public', to: '.', noErrorOnMissing: true },
                    { from: 'src/shortcuts.json', to: 'shortcuts.json' },
                ],
            }),
        ],
        devServer: {
            hot: true,
            port: 3000,
            server: {
                type: 'https',
                options: (() => {
                    if (isProd) return {};
                    try {
                        return require('office-addin-dev-certs').getHttpsServerOptions();
                    } catch {
                        return {};
                    }
                })(),
            },
        },
        resolve: { extensions: ['.js'] },
        devtool: isProd ? false : 'source-map',
    };
};
