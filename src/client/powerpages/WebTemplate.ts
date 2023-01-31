/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 */

import * as vscode from "vscode";
import {
    formatFileName,
    formatFolderName,
    isNullOrEmpty,
} from "./utils/CommonUtils";
import * as nls from "vscode-nls";
import { exec } from "child_process";
import path from "path";
import { yoPath } from "./constants";
nls.config({
    messageFormat: nls.MessageFormat.bundle,
    bundleFormat: nls.BundleFormat.standalone,
})();
const localize: nls.LocalizeFunc = nls.loadMessageBundle();

export const createWebTemplate = (context: vscode.ExtensionContext) => {
    vscode.window
        .showInputBox({
            placeHolder: localize(
                "microsoft-powerapps-portals.webExtension.webtemplate.name",
                "Enter the name of the web template"
            ),
        })
        .then((value) => {
            if (!isNullOrEmpty(value)) {
                // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                const webTemplateFile = formatFileName(value!);
                // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                const webTemplateFolder = formatFolderName(value!);

                const watcher: vscode.FileSystemWatcher =
                    vscode.workspace.createFileSystemWatcher(
                        new vscode.RelativePattern(
                            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                            vscode.workspace.workspaceFolders![0],
                            path.join("web-templates", webTemplateFolder, `${webTemplateFile}.webtemplate.source.html`)
                        ),
                        false,
                        true,
                        true
                    );

                context.subscriptions.push(watcher);
                const portalDir = `${vscode.workspace.workspaceFolders?.[0].uri.fsPath}`;
                const yoSubgenerator = "@microsoft/powerpages:webtemplate";
                const command = `${yoPath} ${yoSubgenerator} "${value}"`;

                vscode.window
                    .withProgress(
                        {
                            location: vscode.ProgressLocation.Notification,
                            title: "Creating web template...",
                        },
                        () => {
                            return new Promise((resolve, reject) => {
                                exec(
                                    command,
                                    { cwd: portalDir },
                                    (error, stderr) => {
                                        if (error) {
                                            vscode.window.showErrorMessage(
                                                error.message
                                            );
                                            reject(error);
                                        } else {
                                            resolve(stderr);
                                        }
                                    }
                                );
                            });
                        }
                    )
                    .then(() => {
                        vscode.window.showInformationMessage(
                            "Web template Created!"
                        );
                    });

                watcher.onDidCreate(async (uri) => {
                    await vscode.window.showTextDocument(uri);
                });
            }
        });
};
