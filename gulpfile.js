/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 */

/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable no-undef */

"use strict";
const util = require('util');
const exec = util.promisify(require('child_process').exec);
const gulp = require('gulp');
const rename = require('gulp-rename');
const eslint = require('gulp-eslint');
const gulpTs = require("gulp-typescript");
const replace = require('gulp-replace');
const mocha = require('gulp-mocha');
const moment = require('moment');
const gulpWebpack = require('webpack-stream');
const webpack = require('webpack');
const vsce = require('@vscode/vsce');
const argv = require('yargs').argv;

const fetch = require('node-fetch');
const fs = require('fs-extra');
const log = require('fancy-log');
const path = require('path');
const pslist = require('ps-list');

const [nodeConfig, webConfig] = require('./webpack.config');
const distdir = path.resolve('./dist');
const outdir = path.resolve('./out');
const packagedir = path.resolve('./package');
const feedPAT = argv.feedPAT || process.env['AZ_DevOps_Read_PAT'];
const isOfficialBuild = argv.isOfficialBuild && argv.isOfficialBuild.toLowerCase() == "true";
const isPreviewBuild = argv.isPreviewBuild && argv.isPreviewBuild.toLowerCase() == "true";

async function clean() {
    (await pslist())
        .filter((info) => info.name.startsWith('pacTelemetryUpload'))
        .forEach(info => {
            log.info(`Terminating: ${info.name} - ${info.pid}...`)
            process.kill(info.pid);
        });
    fs.emptyDirSync(outdir);
    return fs.emptyDir(distdir);
}

function setTelemetryTarget() {
    const telemetryConfigurationSource = isOfficialBuild && !isPreviewBuild
        ? 'src/common/telemetry/telemetryConfigurationProd.ts'
        : 'src/common/telemetry/telemetryConfigurationDev.ts';

    return gulp
        .src(telemetryConfigurationSource)
        .pipe(rename('telemetryConfiguration.ts'))
        .pipe(gulp.dest(path.join('src', 'common', 'telemetry-generated')));
}


function compile() {
    return gulp
        .src('src/**/*.ts')
        .pipe(gulpWebpack(nodeConfig, webpack))
        .pipe(replace("src\\\\client\\\\lib\\\\", "src/client/lib/")) // Hacky fix: vscode-nls-dev/lib/webpack-loader uses Windows style paths when built on Windows, breaking localization on Linux & Mac
        .pipe(gulp.dest(distdir));
}

function compileWeb() {
    return gulp
        .src('src/web/**/*.ts')
        .pipe(gulpWebpack(webConfig, webpack))
        .pipe(replace("src\\\\client\\\\lib\\\\", "src/client/lib/")) // Hacky fix: vscode-nls-dev/lib/webpack-loader uses Windows style paths when built on Windows, breaking localization on Linux & Mac
        .pipe(gulp.dest(path.resolve(`${distdir}/web`)));
}

async function nugetInstall(nugetSource, packageName, version, targetDir) {
    // https://docs.microsoft.com/en-us/nuget/api/package-base-address-resource
    const feeds = {
        'nuget.org': {
            authenticated: false,
            baseUrl: 'https://api.nuget.org/v3-flatcontainer/'
        },
        'CAP_ISVExp_Tools_Daily': {
            authenticated: true,
            // https://dev.azure.com/msazure/One/_packaging?_a=feed&feed=CAP_ISVExp_Tools_Daily
            baseUrl: 'https://pkgs.dev.azure.com/msazure/_packaging/d3fb5788-d047-47f9-9aba-76890f5cecf0/nuget/v3/flat2/'
        },
        'CAP_ISVExp_Tools_Stable': {
            authenticated: true,
            // https://dev.azure.com/msazure/One/_packaging?_a=feed&feed=CAP_ISVExp_Tools_Stable
            baseUrl: 'https://pkgs.dev.azure.com/msazure/_packaging/b0441cf8-0bc8-4fad-b126-841a6184e784/nuget/v3/flat2/'
        },
    }

    const selectedFeed = feeds[nugetSource];
    const baseUrl = selectedFeed.baseUrl;

    packageName = packageName.toLowerCase();
    version = version.toLowerCase();
    const packagePath = `${packageName}/${version}/${packageName}.${version}.nupkg`;

    const nupkgUrl = new URL(packagePath, baseUrl);
    const reqInit = {
        headers: {
            'User-Agent': 'gulpfile-DPX-team/0.1',
            'Accept': '*/*'
        },
        redirect: 'manual'
    };
    if (selectedFeed.authenticated) {
        if (!feedPAT) {
            throw new Error(`nuget feed ${nugetSource} requires authN but neither '--feedToken' argument nor env var 'AZ_DevOps_Read_PAT' was defined!`);
        }
        reqInit.headers['Authorization'] = `Basic ${Buffer.from('PAT:' + feedPAT).toString('base64')}`;
    }

    log.info(`Downloading package: ${nupkgUrl}...`);
    let res = await fetch(nupkgUrl, reqInit);
    if (res.status === 303) {
        const location = res.headers.get('location');
        const url = new URL(location);
        log.info(` ... redirecting to: ${url.origin}${url.pathname}}...`);
        // AzDevOps feeds will redirect to Azure storage with location url w/ SAS token: on 2nd request drop authZ header
        delete reqInit.headers['Authorization'];
        res = await fetch(location, reqInit);
    }
    if (!res.ok) {
        const body = res.body.read();
        throw new Error(`Cannot download ${res.url}, status: ${res.statusText} (${res.status}), body: ${body ? body.toString('ascii') : '<empty>'}`);
    }

    const localNupkg = path.join(targetDir, `${packageName}.${version}.nupkg`);
    fs.ensureDirSync(targetDir);
    return new Promise((resolve, reject) => {
        res.body.pipe(fs.createWriteStream(localNupkg))
            .on('close', () => {
                resolve();
            }).on('error', err => {
                reject(err);
            })
    });
}

function lint() {
    return gulp
        .src(['src/**/*.ts', __filename])
        .pipe(eslint({
            formatter: 'verbose',
            configuration: '.eslintrc.js'
        }))
        .pipe(eslint.format())
        .pipe(eslint.results(results => {
            if (results.warningCount > 0) {
                throw new Error(`Found ${results.warningCount} eslint errors.`)
            }
        }))
        .pipe(eslint.failAfterError());
}

function testUnitTests() {
    return gulp
        .src(
            [
                "src/server/test/unit/**/*.ts",
                "src/client/test/unit/**/*.ts",
                "src/debugger/test/unit/**/*.ts",
                "src/web/client/test/unit/**/*.ts",
            ],
            {
                read: false,
            }
        )
        .pipe(
            mocha({
                require: ["ts-node/register"],
                ui: "bdd",
            })
        );
}

function testWeb() {
    return gulp.src(["src/web/client/test/unit/**/*.ts"], { read: false }).pipe(
        mocha({
            require: ["ts-node/register"],
            ui: "bdd",
        })
    );
}

// unit tests without special test runner
const test = gulp.series(testUnitTests, testWeb);

/**
 * Compiles the integration tests and transpiles the results to /out
 */
function compileIntegrationTests() {
    const tsProject = gulpTs.createProject("tsconfig.json", {
        // to test puppeteer we need "dom".
        // since "dom" overlaps with "webworker" we need to overwrite the lib property.
        // This is a known ts issue (bot being able to have both webworker and dom): https://github.com/microsoft/TypeScript/issues/20595
        lib: ["es2019", "dom", "dom.iterable"],
    });
    return gulp.src(["src/**/*.ts"]).pipe(tsProject()).pipe(gulp.dest("out"));
}

/**
 * Tests the debugger integration tests after transpiling the source files to /out
 */
const testDebugger = gulp.series(compileIntegrationTests, async () => {
    const testRunner = require("./out/debugger/test/runTest");
    await testRunner.main();
});

// tests that require vscode-electron (which requires a display or xvfb)
const testInt = gulp.series(testDebugger);

/**
 * Tests the debugger integration tests after transpiling the source files to /out
 */
const testWebIntegration = gulp.series(compileIntegrationTests, async () => {
    const testRunner = require("./out/web/client/test/runTest");
    await testRunner.main();
});

// tests that require vscode-electron (which requires a display or xvfb)
const testWebInt = gulp.series(testWebIntegration);

/**
 * Tests the power-pages integration tests after transpiling the source files to /out
 */
const testDesktopIntegration = gulp.series(compileIntegrationTests, async () => {
    const testRunner = require("./out/client/test/runTest");
    await testRunner.main();
});

// tests that require vscode-electron (which requires a display or xvfb)
const testDesktopInt = gulp.series(testDesktopIntegration);

async function packageVsix() {
    const standardHeader = '# Power Platform Extension';
    const previewHeader = '# Power Platform Tools [PREVIEW]\n\n## This extension is used for internal testing against targets such as vscode.dev which require Marketplace published extensions, and is not supported.';
    const standardPackageOptions = {
        name: 'powerplatform-vscode',
        displayName: 'Power Platform Tools',
        description: 'Tooling to create Power Platform solutions & packages, manage Power Platform environments and edit Power Apps Portals',
        readmeHeader: standardHeader,
        readmeReplacementTarget: previewHeader,
    };
    const previewPackageOptions = {
        name: 'powerplatform-vscode-preview',
        displayName: 'Power Platform Tools [PREVIEW]',
        description: 'Unsupported extension for testing Power Platform Tools',
        readmeHeader: previewHeader,
        readmeReplacementTarget: standardHeader,
     };

    const setPackageInfo = async function(pkgOptions) {
        await npm(['pkg', 'set', `name=${pkgOptions.name}`]);
        await npm(['pkg', 'set', `displayName="${pkgOptions.displayName}"`]);
        await npm(['pkg', 'set', `description="${pkgOptions.description}"`]);

        gulp.src('README.md')
            .pipe(replace(pkgOptions.readmeReplacementTarget, pkgOptions.readmeHeader))
            .pipe(gulp.dest('./'));
    }

    await setPackageInfo(isPreviewBuild ? previewPackageOptions : standardPackageOptions);

    await vsce.createVSIX({
        packagePath: packagedir,
        preRelease: isPreviewBuild,
    });

    // Reset to non-preview settings to prevent polluting git diffs
    if (isPreviewBuild) {
        await setPackageInfo(standardPackageOptions);
    }
}

async function git(args) {
    args.unshift('git');
    const { stdout, stderr } = await exec(args.join(' '));
    return { stdout: stdout, stderr: stderr };
}

async function npm(args) {
    args.unshift('npm');
    const {stdout, stderr } = await exec(args.join(' '));
    return {stdout: stdout, stderr: stderr};
}

async function npx(args) {
    args.unshift('npx');
    const {stdout, stderr } = await exec(args.join(' '));
    return {stdout: stdout, stderr: stderr};
}

async function setGitAuthN() {
    const repoUrl = 'https://github.com';
    const repoToken = argv.repoToken;
    if (!repoToken) {
        throw new Error(`Must specify parameter --repoToken with read and push rights to ${repoUrl}!`);
    }
    const bearer = `AUTHORIZATION: basic ${Buffer.from(`PAT:${repoToken}`).toString('base64')}`;
    await git(['config', '--local', `http.${repoUrl}/.extraheader`, `"${bearer}"`]);
    await git(['config', '--local', 'user.email', 'capisvaatdev@microsoft.com']);
    await git(['config', '--local', 'user.name', '"DPT Tools Dev Team"']);
}

async function snapshot() {
    const targetBranch = argv.targetBranch || 'release/daily';
    const sourceSpecParam = argv.sourceSpec;

    const tmpRepo = path.resolve('./out/tmpRepo');
    fs.emptyDirSync(tmpRepo);

    const repoUrl = (await git(['remote', 'get-url', '--all', 'origin'])).stdout.trim();
    log.info(`snapshot: remote repoUrl: ${repoUrl}`);
    const orgDir = process.cwd();
    process.chdir(tmpRepo);
    try {
        await git(['init']);
        await git(['remote', 'add', 'origin', repoUrl]);
        await setGitAuthN();
        await git(['fetch', 'origin']);
        const remotes = (await git(['remote', 'show', 'origin'])).stdout;
        const head = remotes
            .split('\n')
            .map(line => {
                const branch = line.match(/HEAD branch:\s*(\S+)/);
                if (branch && branch.length >= 2) {
                    return branch[1];
                }
            })
            .filter(b => !!b);
        if (!head || head.length < 1 || head.length > 1 || !head[0]) {
            throw new Error(`Cannot determine HEAD from remote: ${repoUrl}`);
        }
        const headBranch = head[0];
        if (headBranch == targetBranch) {
            throw new Error(`Cannot snapshot into default HEAD branch: ${headBranch}`);
        }
        const sourceSpec = sourceSpecParam || `origin/${headBranch}`;
        log.info(`  > snap shotting '${sourceSpec}' into branch: ${targetBranch}...`);
        await git(['checkout', headBranch]);
        const snapshotTag = `snapshot-${targetBranch.replace('/', '_').replace(' ', '_')}-${moment.utc().format('YYMMDD[Z]HHmmss')}`;
        // await git(['tag', snapshotTag, sourceSpec]);
        await git(['checkout', '--force', '-B', targetBranch]);
        const resetMsg = (await git(['reset', '--hard', `"${sourceSpec}"`])).stdout.trim();
        log.info(`  > snapshot (${snapshotTag}): ${resetMsg}`);
        log.info(`  > pushing snapshot branch '${targetBranch} to origin...`);
        const pushMsg = (await git(['push', '--force', '--tags', 'origin', targetBranch])).stderr.trim();
        log.info(`  > ${pushMsg}`)
    }
    finally {
        process.chdir(orgDir);
    }
}

const feedName = 'CAP_ISVExp_Tools_Stable';
const cliVersion = '1.21.13';

const recompile = gulp.series(
    clean,
    async () => nugetInstall(feedName, 'Microsoft.PowerApps.CLI', cliVersion, path.resolve(distdir, 'pac')),
    async () => nugetInstall(feedName, 'Microsoft.PowerApps.CLI.Core.osx-x64', cliVersion, path.resolve(distdir, 'pac')),
    async () => nugetInstall(feedName, 'Microsoft.PowerApps.CLI.Core.linux-x64', cliVersion, path.resolve(distdir, 'pac')),
    translationsExport,
    translationsImport,
    setTelemetryTarget,
    compile,
    compileWeb
);

const dist = gulp.series(
    recompile,
    packageVsix,
    lint,
    test
);
// Extract all the localizable strings from TS and package.nls.json, and package into
// an XLF for the localization team
async function translationsExport() {
    await npx(["@vscode/l10n-dev", "export", "--outDir", "./l10n", "./src"]);
    await npx(["@vscode/l10n-dev", "generate-xlf",
        "./package.nls.json", "./l10n/bundle.l10n.json",
        "--outFile", "./loc/translations-export/vscode-powerplatform.xlf"]);
    return gulp.src('./loc/translations-export/vscode-powerplatform.xlf')
        .pipe(replace("&apos;", "'"))
        .pipe(replace("&#10;", "\n"))
        .pipe(gulp.dest('./loc/translations-export/'));
}

// const languages = [
//     //{ id: "bg", folderName: "bul" },
//     //{ id: "hu", folderName: "hun" },
//     //{ id: "pl", folderName: "plk" },
//     { id: "cs", folderName: "csy" },
//     { id: "de", folderName: "deu" },
//     { id: "es", folderName: "esn" },
//     { id: "fr", folderName: "fra" },
//     { id: "it", folderName: "ita" },
//     { id: "ja", folderName: "jpn" },
//     { id: "ko", folderName: "kor" },
//     { id: "pt-BR", folderName: "ptb" },
//     { id: "ru", folderName: "rus" },
//     { id: "tr", folderName: "trk" },
//     { id: "zh-CN", folderName: "chs" },
//     { id: "zh-TW", folderName: "cht" },
// ];

async function translationsImport() {
    await npx(["@vscode/l10n-dev", "import-xlf", "./loc/translations-import/*.xlf", "--outDir", "./l10n"]);

    // `@vscode/l10n-dev import-xlf` places both the package.nls.*.json and bundle.l10n.*.json files in the
    // same directory, but the package.nls.*.json need to reside at the repo root next to package.json
    gulp.src('./l10n/package.nls.*.json')
        .pipe(replace("\\r\\n", "\\n"))
        .pipe(replace("\\\\n", "\\n"))
        .pipe(gulp.dest('./'));

    // Fix up changes from the XLF Export / Import process that cause lookup misses
    return gulp.src('./l10n/bundle.l10n.*.json')
        .pipe(replace("\\r\\n", "\\n"))
        .pipe(replace("\\\\n", "\\n"))
        .pipe(replace("&apos;", "'"))
        .pipe(gulp.dest('./l10n'));
}

exports.clean = clean;
exports.compile = compile;
exports.compileWeb = compileWeb;
exports.recompile = recompile;
exports.snapshot = snapshot;
exports.lint = lint;
exports.test = test;
exports.testWeb = testWeb;
exports.compileIntegrationTests = compileIntegrationTests;
exports.testInt = testInt;
exports.testWebInt = testWebInt;
exports.testDesktopInt = testDesktopInt;
exports.package = packageVsix;
exports.ci = dist;
exports.dist = dist;
exports.translationsExport = translationsExport;
exports.translationsImport = translationsImport;
exports.setGitAuthN = setGitAuthN;
exports.default = compile;
