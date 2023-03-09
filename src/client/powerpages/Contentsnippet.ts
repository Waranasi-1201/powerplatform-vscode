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
import { QuickPickItem } from "vscode";
import { MultiStepInput } from "./utils/MultiStepInput";
import { exec } from "child_process";
import path from "path";
import { statSync } from "fs";

export const createContentSnippet = async (
    context: vscode.ExtensionContext,
    selectedWorkspaceFolder: string | undefined,
    yoPath: string | null
) => {
    if (!selectedWorkspaceFolder) {
        throw new Error("Root directory not found");
    }
    const { contentSnippetName, contentSnippetType } =
        await getContentSnippetInputs(selectedWorkspaceFolder);

    if (!isNullOrEmpty(contentSnippetName)) {
        const folder = formatFolderName(contentSnippetName);
        const file = formatFileName(contentSnippetName);

        const watcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(
                // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                selectedWorkspaceFolder!,
                path.join(
                    "content-snippets",
                    folder,
                    `${file}.*.contentsnippet.yml`
                )
            ),
            false,
            true,
            true
        );

        context.subscriptions.push(watcher);
        const portalDir = selectedWorkspaceFolder;
        const yoContentSnippetGenerator = "@microsoft/powerpages:contentsnippet";
        const command = `"${yoPath}" ${yoContentSnippetGenerator} "${contentSnippetName}" "${contentSnippetType}"`;

        vscode.window
            .withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: "Creating Content Snippet...",
                },
                () => {
                    return new Promise((resolve, reject) => {
                        exec(command, { cwd: portalDir }, (error, stderr) => {
                            if (error) {
                                vscode.window.showErrorMessage(error.message);
                                reject(error);
                            } else {
                                resolve(stderr);
                            }
                        });
                    });
                }
            )
            .then(() => {
                vscode.window.showInformationMessage(
                    "Content Snippet Created!"
                );
            });

        watcher.onDidCreate(async (uri) => {
            await vscode.window.showTextDocument(uri);
        });
    }
};

async function getContentSnippetInputs(selectedWorkspaceFolder: string) {
    const contentSnippetTypes: QuickPickItem[] = ["html", "text"].map(
        (label) => ({ label })
    );

    interface State {
        title: string;
        step: number;
        totalSteps: number;
        contentSnippetType: QuickPickItem | string;
        contentSnippetName: string;
    }

    async function collectInputs() {
        const state = {} as Partial<State>;
        await MultiStepInput.run((input) => inputName(input, state));
        return state as State;
    }

    const title =
        "New Content Snippet"
    ;

    async function inputName(input: MultiStepInput, state: Partial<State>) {
        state.contentSnippetName = await input.showInputBox({
            title,
            step: 1,
            totalSteps: 2,
            value: state.contentSnippetName || "",
            placeholder:
                "Add content snippet name (name should be unique)"
            ,
            validate: validateNameIsUnique,
        });
        return (input: MultiStepInput) => pickType(input, state);
    }

    async function pickType(input: MultiStepInput, state: Partial<State>) {
        const pick = await input.showQuickPick({
            title,
            step: 2,
            totalSteps: 2,
            placeholder:
                "Select Type"
            ,
            items: contentSnippetTypes,
            activeItem:
                typeof state.contentSnippetType !== "string"
                    ? state.contentSnippetType
                    : undefined,
        });

        state.contentSnippetType = pick.label;
    }

    async function validateNameIsUnique(
        name: string
    ): Promise<string | undefined> {
        const folder = formatFolderName(name);
        const filePath = path.join(
            selectedWorkspaceFolder,
            "content-snippets",
            folder
        );
        try {
            const stat = statSync(filePath);
            if (stat) {
                return "A content snippet with the same name already exists. Please enter a different name.";
            }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        } catch (error: any) {
            if (error.code === "ENOENT") {
                return undefined;
            }
        }
    }

    const state = await collectInputs();
    return state;
}
