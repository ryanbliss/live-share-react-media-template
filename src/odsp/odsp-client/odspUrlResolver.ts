/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { IRequest } from "@fluidframework/core-interfaces";
import { IUrlResolver, IResolvedUrl } from "@fluidframework/driver-definitions";
import {
    IOdspAuthRequestInfo,
    getDriveItemByRootFileName,
} from "@fluidframework/odsp-doclib-utils";
import {
    OdspDriverUrlResolver,
    createOdspUrl,
    createOdspCreateContainerRequest,
} from "@fluidframework/odsp-driver";

export class OdspUrlResolver implements IUrlResolver {
    private readonly driverUrlResolver = new OdspDriverUrlResolver();

    constructor(
        private readonly server: string,
        private readonly authRequestInfo: IOdspAuthRequestInfo
    ) {}

    public async resolve(request: IRequest): Promise<IResolvedUrl> {
        try {
            const resolvedUrl = await this.driverUrlResolver.resolve(request);
            console.log(
                "odspUrlResolver::resolve: starting url",
                request.url,
                "\nending url:",
                resolvedUrl
            );
            return resolvedUrl;
        } catch (error) {
            console.error(error);
        }

        console.log("odsp url resolver: ", request);
        const url = new URL(request.url);

        const fullPath = url.pathname.substr(1);
        const documentId = fullPath.split("/")[0];
        const dataStorePath = fullPath.slice(documentId.length + 1);
        const filePath = this.formFilePath(documentId);

        const { driveId, itemId } = await getDriveItemByRootFileName(
            this.server,
            "",
            filePath,
            this.authRequestInfo,
            true
        );

        const odspUrl = createOdspUrl({
            siteUrl: `https://${this.server}`,
            driveId,
            itemId,
            dataStorePath,
        });

        return this.driverUrlResolver.resolve({
            url: odspUrl,
            headers: request.headers,
        });
    }

    private formFilePath(documentId: string): string {
        const encoded = encodeURIComponent(`${documentId}.fluid`);
        return `/r11s/${encoded}`;
    }

    public async getAbsoluteUrl(
        resolvedUrl: IResolvedUrl,
        relativeUrl: string
    ): Promise<string> {
        const absoluteUrl = await this.driverUrlResolver.getAbsoluteUrl(
            resolvedUrl,
            relativeUrl
        );
        console.log(
            "odspUrlResolver::getAbsoluteUrl: getting absolute url",
            absoluteUrl
        );
        return absoluteUrl;
    }

    public async createCreateNewRequest(fileName: string): Promise<IRequest> {
        console.log(
            "odspUrlResolver::createCreateNewRequest: creating new request for",
            fileName
        );
        const filePath = "/r11s/";
        const driveItem = await getDriveItemByRootFileName(
            this.server,
            "",
            filePath,
            this.authRequestInfo,
            false
        );
        return createOdspCreateContainerRequest(
            `https://${this.server}`,
            driveItem.driveId,
            filePath,
            `${fileName}.fluid`
        );
    }
}
