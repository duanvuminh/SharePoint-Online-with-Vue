'use strict';

// check if gulp dist was called
if (process.argv.indexOf('dist') !== -1) {
    // add ship options to command call
    process.argv.push('--ship');
}

const path = require('path');
const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const gulpSequence = require('gulp-sequence');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Create clean distrubution package
gulp.task('dist', gulpSequence('clean', 'bundle', 'package-solution'));
// Create clean development package
gulp.task('dev', gulpSequence('clean', 'bundle', 'package-solution'));



/**
 * Webpack Bundle Anlayzer
 * Reference and gulp task
 */
const bundleAnalyzer = require('webpack-bundle-analyzer');

build.configureWebpack.mergeConfig({

    additionalConfiguration: (generatedConfiguration) => {
        const lastDirName = path.basename(__dirname);
        const dropPath = path.join(__dirname, 'temp', 'stats');
        generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin({
            openAnalyzer: false,
            analyzerMode: 'static',
            reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
            generateStatsFile: true,
            statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
            logLevel: 'error'
        }));

        return generatedConfiguration;
    }

});


/**
 * Custom Framework Specific gulp tasks
 */
// Merge custom loader to web pack configuration

const VueLoaderPlugin = require('vue-loader/lib/plugin');
const vuePlugin = new VueLoaderPlugin();
const themedStyleLoader = require.resolve('@microsoft/loader-load-themed-styles');


const ForkTsCheckerWebpackPlugin = require('fork-ts-checker-webpack-plugin');
    const forkTsPlugin = new ForkTsCheckerWebpackPlugin({
        vue: true,
        tslint: true,
        formatter: 'codeframe',
        checkSyntacticErrors: false,
        watch: '/src/**/*.vue'
    });


    build.configureWebpack.mergeConfig({

        additionalConfiguration: (generatedConfiguration) => {

            const loadersConfigs = [{
                test: /\.vue$/, // vue
                use: [{
                    loader: 'vue-loader'
                }]
            }, {
                resourceQuery: /vue&type=script/, // typescript
                loader: 'ts-loader',
                options: {
                    appendTsSuffixTo: [/\.vue$/],
                    transpileOnly: true
                }
            }, {
                resourceQuery: /vue&type=style.*&lang=scss/, // scss
                use: [{
                    loader: themedStyleLoader,
                    options: {
                        async: true
                    }
                },
                {
                    loader: 'css-loader',
                    options: {
                        modules: true,
                        localIdentName: '[local]_[sha1:hash:hex:8]'
                    }
                },
                    'sass-loader'
                ]
            }];

            generatedConfiguration.plugins.push(vuePlugin);
        
                generatedConfiguration.plugins.push(forkTsPlugin);
        
                generatedConfiguration.module.rules.push(...loadersConfigs);

            return generatedConfiguration;

        }

    });

// register custom watch for Vue.JS files
// copy of '.vue' files will be handled by 'copy-static-assets.json'
gulp.watch('./src/**/*.vue', event => {
    // copy empty index.ts onto itself to launch build procees
    gulp.src('./src/index.ts')
        .pipe(gulp.dest('./src/'));

});


build.initialize(gulp);
