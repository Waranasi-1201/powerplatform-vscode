/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 */

import * as path from "path";
import * as vscode from "vscode";
import {
    getRequestURL,
    pathHasEntityFolderName,
} from "../utilities/urlBuilderUtil";
import {
    httpMethod,
    PORTALS_URI_SCHEME,
    queryParameters,
} from "../common/constants";
import { FileData } from "../context/fileData";
import WebExtensionContext from "../WebExtensionContext";
import { fetchDataFromDataverseAndUpdateVFS } from "./remoteFetchProvider";
import { saveData } from "./remoteSaveProvider";
import { ERRORS } from "../common/errorHandler";
import { telemetryEventNames } from "../telemetry/constants";
import { getEntity } from "../utilities/schemaHelperUtil";
import { folderExportType, schemaEntityKey } from "../schema/constants";

export class File implements vscode.FileStat {
    type: vscode.FileType;
    ctime: number;
    mtime: number;
    size: number;

    name: string;
    data: Uint8Array;

    constructor(name: string) {
        this.type = vscode.FileType.File;
        this.ctime = Date.now();
        this.mtime = Date.now();
        this.size = 0;
        this.name = name;
        this.data = new Uint8Array();
    }
}

export class Directory implements vscode.FileStat {
    type: vscode.FileType;
    ctime: number;
    mtime: number;
    size: number;

    name: string;
    entries: Map<string, File | Directory>;

    constructor(name: string) {
        this.type = vscode.FileType.Directory;
        this.ctime = Date.now();
        this.mtime = Date.now();
        this.size = 0;
        this.name = name;
        this.entries = new Map();
    }
}

export type Entry = File | Directory;

export class PortalsFS implements vscode.FileSystemProvider {
    root = new Directory("");

    // --- manage file metadata

    async stat(uri: vscode.Uri): Promise<vscode.FileStat> {
        return await this._lookup(uri, false);
    }

    async readDirectory(uri: vscode.Uri): Promise<[string, vscode.FileType][]> {
        const result: [string, vscode.FileType][] = [];
        try {
            const entry = await this._lookupAsDirectory(uri, false);
            for (const [name, child] of entry.entries) {
                result.push([name, child.type]);
            }
        } catch (error) {
            const castedError = error as vscode.FileSystemError;

            if (castedError.code === vscode.FileSystemError.FileNotFound.name) {
                if (
                    WebExtensionContext.isContextSet &&
                    uri.toString().toLowerCase() ===
                        WebExtensionContext.rootDirectory
                            .toString()
                            .toLowerCase()
                ) {
                    WebExtensionContext.telemetry.sendInfoTelemetry(
                        telemetryEventNames.WEB_EXTENSION_FETCH_DIRECTORY_TRIGGERED
                    );
                    await this._loadFromDataverseToVFS();
                }
            }
        }
        return result;
    }

    // --- manage file contents

    async readFile(uri: vscode.Uri): Promise<Uint8Array> {
        try {
            const data = await this._lookupAsFile(uri, false);
            return data.data;
        } catch (error) {
            const castedError = error as vscode.FileSystemError;

            if (castedError.code === vscode.FileSystemError.FileNotFound.name) {
                if (
                    WebExtensionContext.isContextSet &&
                    uri
                        .toString()
                        .includes(WebExtensionContext.rootDirectory.toString())
                ) {
                    if (pathHasEntityFolderName(uri.toString())) {
                        WebExtensionContext.telemetry.sendInfoTelemetry(
                            telemetryEventNames.WEB_EXTENSION_FETCH_FILE_TRIGGERED
                        );
                        await this._loadFromDataverseToVFS();
                        const data = await this._lookupAsFile(uri, false);
                        return data.data;
                    }
                }
            }
        }
        return new Uint8Array();
    }

    async writeFile(
        uri: vscode.Uri,
        content: Uint8Array,
        options: { create: boolean; overwrite: boolean }
    ): Promise<void> {
        const basename = path.posix.basename(uri.path);
        const parent = await this._lookupParentDirectory(uri);
        let entry = parent.entries.get(basename);
        if (entry instanceof Directory) {
            throw vscode.FileSystemError.FileIsADirectory(uri);
        }
        if (!entry && !options.create) {
            throw vscode.FileSystemError.FileNotFound();
        }
        if (entry && options.create && !options.overwrite) {
            throw vscode.FileSystemError.FileExists(uri);
        }
        if (!entry) {
            entry = new File(basename);
            parent.entries.set(basename, entry);
            this._fireSoon({ type: vscode.FileChangeType.Created, uri });
        } else if (
            WebExtensionContext.fileDataMap.getFileMap.get(uri.fsPath)
                ?.hasDirtyChanges
        ) {
            // Save data to dataverse
            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    cancellable: true,
                    title: vscode.l10n.t("Saving your file ..."),
                },
                async () => {
                    WebExtensionContext.telemetry.sendInfoTelemetry(
                        telemetryEventNames.WEB_EXTENSION_SAVE_FILE_TRIGGERED
                    );
                    await this._saveFileToDataverseFromVFS(uri);
                }
            );
        }

        entry.mtime = Date.now();
        entry.size = content.byteLength;
        entry.data = content;

        this._fireSoon({ type: vscode.FileChangeType.Changed, uri });
    }

    // --- manage files/folders
    async createDirectory(uri: vscode.Uri): Promise<void> {
        // Do silent lookup to check for existing entry
        const entry = await this._lookup(uri, true);
        if (!entry) {
            const basename = path.posix.basename(uri.path);
            const dirname = uri.with({ path: path.posix.dirname(uri.path) });
            const parent = await this._lookupAsDirectory(dirname, false);

            const entry = new Directory(basename);
            parent.entries.set(entry.name, entry);
            parent.mtime = Date.now();
            parent.size += 1;
            this._fireSoon(
                { type: vscode.FileChangeType.Changed, uri: dirname },
                { type: vscode.FileChangeType.Created, uri }
            );
        }
    }

    async rename(): Promise<void> {
        throw new Error("Method not implemented.");
    }

    async delete(): Promise<void> {
        throw new Error("Method not implemented.");
    }

    // --- lookup

    private async _lookup(uri: vscode.Uri, silent: false): Promise<Entry>;
    private async _lookup(
        uri: vscode.Uri,
        silent: boolean
    ): Promise<Entry | undefined>;
    private async _lookup(
        uri: vscode.Uri,
        silent: boolean
    ): Promise<Entry | undefined> {
        const parts = uri.path.split("/");
        let entry: Entry = this.root;
        for (const part of parts) {
            if (!part) {
                continue;
            }
            let child: Entry | undefined;
            if (entry instanceof Directory) {
                child = entry.entries.get(part);
            }
            if (!child) {
                if (!silent) {
                    throw vscode.FileSystemError.FileNotFound();
                } else {
                    return undefined;
                }
            }
            entry = child;
        }
        return entry;
    }

    private async _lookupAsDirectory(
        uri: vscode.Uri,
        silent: boolean
    ): Promise<Directory> {
        const entry = await this._lookup(uri, silent);
        if (entry instanceof Directory) {
            return entry;
        }
        throw vscode.FileSystemError.FileNotADirectory(uri);
    }

    private async _lookupAsFile(
        uri: vscode.Uri,
        silent: boolean
    ): Promise<File> {
        const entry = await this._lookup(uri, silent);
        if (entry instanceof File) {
            return entry;
        }
        throw vscode.FileSystemError.FileIsADirectory(uri);
    }

    private async _lookupParentDirectory(uri: vscode.Uri): Promise<Directory> {
        const dirname = uri.with({ path: path.posix.dirname(uri.path) });
        return await this._lookupAsDirectory(dirname, false);
    }

    // --- manage file events

    private _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    private _bufferedEvents: vscode.FileChangeEvent[] = [];
    private _fireSoonHandle?: NodeJS.Timeout;

    readonly onDidChangeFile: vscode.Event<vscode.FileChangeEvent[]> =
        this._emitter.event;

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    watch(_resource: vscode.Uri): vscode.Disposable {
        // ignore, fires for all changes...
        // eslint-disable-next-line @typescript-eslint/no-empty-function
        return new vscode.Disposable(() => {});
    }

    private _fireSoon(...events: vscode.FileChangeEvent[]): void {
        this._bufferedEvents.push(...events);

        if (this._fireSoonHandle) {
            clearTimeout(this._fireSoonHandle);
        }
        this._fireSoonHandle = setTimeout(() => {
            this._emitter.fire(this._bufferedEvents);
            this._bufferedEvents.length = 0;
        }, 5);
    }

    // --- VFS calls
    private async createFileSystem(portalFolderName: string) {
        if (portalFolderName.length === 0) {
            WebExtensionContext.telemetry.sendErrorTelemetry(
                telemetryEventNames.WEB_EXTENSION_EMPTY_PORTAL_FOLDER_NAME
            );
            throw new Error(ERRORS.PORTAL_FOLDER_NAME_EMPTY);
        }

        WebExtensionContext.telemetry.sendInfoTelemetry(
            telemetryEventNames.WEB_EXTENSION_CREATE_ROOT_FOLDER
        );

        await this.createDirectory(
            vscode.Uri.parse(
                `${PORTALS_URI_SCHEME}:/${portalFolderName}/`,
                true
            )
        );
        vscode.workspace.updateWorkspaceFolders(
            vscode.workspace.workspaceFolders
                ? vscode.workspace.workspaceFolders.length
                : 0,
            null,
            {
                uri: vscode.Uri.parse(
                    `${PORTALS_URI_SCHEME}:/${portalFolderName}/`
                ),
                name: portalFolderName,
            }
        );

        await this.createEntityFolder(portalFolderName);
    }

    private async createEntityFolder(portalFolderName: string) {
        const entityDetails = getEntity(WebExtensionContext.defaultEntityType);
        const exportType = entityDetails?.get(schemaEntityKey.EXPORT_TYPE);
        const subUri = entityDetails?.get(schemaEntityKey.FILE_FOLDER_NAME);

        if (subUri?.length === 0) {
            throw new Error(ERRORS.SUBURI_EMPTY);
        }
        WebExtensionContext.telemetry.sendInfoTelemetry(
            telemetryEventNames.WEB_EXTENSION_CREATE_ENTITY_FOLDER
        );

        let filePathInPortalFS = "";
        if (
            exportType &&
            (exportType === folderExportType.SubFolders ||
                exportType === folderExportType.SingleFolder)
        ) {
            filePathInPortalFS = `${PORTALS_URI_SCHEME}:/${portalFolderName}/${subUri}/`;
            await this.createDirectory(
                vscode.Uri.parse(filePathInPortalFS, true)
            );
        }
    }

    // --- Dataverse calls

    private async _loadFromDataverseToVFS() {
        await WebExtensionContext.authenticateAndUpdateDataverseProperties();
        await this.createFileSystem(
            WebExtensionContext.urlParametersMap.get(
                queryParameters.WEBSITE_NAME
            ) as string
        );

        await fetchDataFromDataverseAndUpdateVFS(this);
    }

    private async _saveFileToDataverseFromVFS(uri: vscode.Uri) {
        const dataMap: Map<string, FileData> =
            WebExtensionContext.fileDataMap.getFileMap;
        const dataverseOrgUrl = WebExtensionContext.urlParametersMap.get(
            queryParameters.ORG_URL
        ) as string;

        await WebExtensionContext.reAuthenticate();
        const patchRequestUrl = getRequestURL(
            dataverseOrgUrl,
            dataMap.get(uri.fsPath)?.entityName as string,
            dataMap.get(uri.fsPath)?.entityId as string,
            httpMethod.PATCH,
            true
        );

        await saveData(
            WebExtensionContext.dataverseAccessToken,
            patchRequestUrl,
            uri
        );
    }
}
