const path = require('path');
const webpack = require('webpack');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const ESLintPlugin = require('eslint-webpack-plugin');
const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;

module.exports = (env) => {
    //where react is getting is elements to load into the DOM

    return {
        entry: {
            main: { import: './Components/main.jsx' },
        },
        devtool: 'source-map',
        module: {
            rules: [
                {
                    //react rules
                    test: /\.(js|jsx)$/,
                    exclude: /(node_modules)/,
                    use: ['babel-loader'],
                },
                {
                    test: /\.s[ac]ss$/i,
                    //test: /\.(scss)$/,
                    use: [
                        'style-loader', // inject CSS to page
                        'css-loader', // translates CSS into CommonJS modules
                        // Run postcss actions
                        {
                            loader: 'postcss-loader',
                            options: {
                                // `postcssOptions` is needed for postcss 8.x;
                                postcssOptions: {
                                    // postcss plugins, can be exported to postcss.config.js
                                    plugins: [
                                        function () {
                                            return [require('autoprefixer')];
                                        },
                                        function () {
                                            return [require('cssnano')];
                                        },
                                    ],
                                },
                            },
                        },
                        {
                            loader: 'sass-loader', // compiles Sass to CSS
                            options: {
                                // Prefer `dart-sass`
                                implementation: require('sass'),
                            },
                        },
                    ],
                },
                {
                    test: /\.(svg|eot|woff|woff2|ttf)$/,
                    use: ['file-loader'],
                },
            ],
        },
        resolve: { extensions: ['*', '.js', '.jsx'] },
        output: {
            path: path.resolve(__dirname, 'wwwroot/public'),
            publicPath: '/public/',
            filename: '[name].bundle.js',
            clean: true,
        },
        optimization: {
            splitChunks: {
                chunks: 'all',
            },
        },
        plugins: [
            new CleanWebpackPlugin(),
            new ESLintPlugin({
                extensions: ['.js', '.jsx'],
            }),
            new BundleAnalyzerPlugin({
                analyzerMode: 'static',
                openAnalyzer: false,
            }),
        ],
    };
};
