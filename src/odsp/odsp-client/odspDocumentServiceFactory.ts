/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { ISummaryTree } from "@fluidframework/protocol-definitions";
import {
    OdspResourceTokenFetchOptions,
    TokenFetcher,
    IPersistedCache,
    HostStoragePolicy,
    IOdspUrlParts,
    ShareLinkTypes,
} from "@fluidframework/odsp-driver-definitions";
import { getDocAttributesFromProtocolSummary } from "@fluidframework/driver-utils";
import {
    IDocumentService,
    IResolvedUrl,
} from "@fluidframework/driver-definitions";
import { PerformanceEvent } from "@fluidframework/telemetry-utils";
import {
    IExistingFileInfo,
    INewFileInfo,
    getOdspResolvedUrl,
    isNewFileInfo,
    createOdspLogger,
    toInstrumentedOdspTokenFetcher,
} from "./odspUtils";
import { isCombinedAppAndProtocolSummary } from "./driverUtils";
import {
    LocalPersistentCache,
    NonPersistentCache,
} from "@fluidframework/odsp-driver/dist/odspCache";
import { createOdspCacheAndTracker } from "./epochTracker";
import { OdspDocumentServiceFactoryCore } from "@fluidframework/odsp-driver";
import { getSocketIo } from "./getSocketIo";
import { createNewFluidFile } from "./createFile";
import { createNewContainerOnExistingFile } from "./createNewContainerOnExistingFile";

/**
 * Factory for creating the sharepoint document service. Use this if you want to
 * use the sharepoint implementation.
 * @public
 */
export class PolyfillOdspDocumentServiceFactory extends OdspDocumentServiceFactoryCore {
    protected _getStorageToken: TokenFetcher<OdspResourceTokenFetchOptions>;
    protected _getWebsocketToken:
        | TokenFetcher<OdspResourceTokenFetchOptions>
        | undefined;
    protected _hostPolicy: HostStoragePolicy;
    protected _persistedCache: IPersistedCache;
    private readonly _nonPersistentCache = new NonPersistentCache();
    constructor(
        getStorageToken: TokenFetcher<OdspResourceTokenFetchOptions>,
        getWebsocketToken:
            | TokenFetcher<OdspResourceTokenFetchOptions>
            | undefined,
        persistedCache: IPersistedCache = new LocalPersistentCache(),
        hostPolicy: HostStoragePolicy = {}
    ) {
        super(
            getStorageToken,
            getWebsocketToken,
            async () => getSocketIo(),
            persistedCache,
            hostPolicy
        );
        this._getStorageToken = getStorageToken;
        this._getWebsocketToken = getWebsocketToken;
        this._persistedCache = persistedCache;
        this._hostPolicy = hostPolicy;
        // Set enableRedeemFallback by default as true.
        this._hostPolicy.enableRedeemFallback =
            this._hostPolicy.enableRedeemFallback ?? true;
        this._hostPolicy.sessionOptions = {
            forceAccessTokenViaAuthorizationHeader: true,
            ...this._hostPolicy.sessionOptions,
        };
    }

    public override async createContainer(
        createNewSummary: ISummaryTree | undefined,
        createNewResolvedUrl: IResolvedUrl,
        // Made any type to resolve `ITelemetryBaseLogger` not being exported
        logger?: any,
        // logger?: ITelemetryBaseLogger,
        clientIsSummarizer?: boolean
    ): Promise<IDocumentService> {
        let odspResolvedUrl = getOdspResolvedUrl(createNewResolvedUrl);
        const resolvedUrlData: IOdspUrlParts = {
            siteUrl: odspResolvedUrl.siteUrl,
            driveId: odspResolvedUrl.driveId,
            itemId: odspResolvedUrl.itemId,
        };

        let fileInfo: INewFileInfo | IExistingFileInfo;
        let createShareLinkParam: ShareLinkTypes | undefined;
        if (odspResolvedUrl.itemId) {
            fileInfo = {
                type: "Existing",
                driveId: odspResolvedUrl.driveId,
                siteUrl: odspResolvedUrl.siteUrl,
                itemId: odspResolvedUrl.itemId,
            };
        } else if (odspResolvedUrl.fileName) {
            const [, queryString] = odspResolvedUrl.url.split("?");
            const searchParams = new URLSearchParams(queryString);
            const filePath = searchParams.get("path");
            if (filePath === undefined || filePath === null) {
                throw new Error("File path should be provided!!");
            }
            createShareLinkParam = getSharingLinkParams(
                this._hostPolicy,
                searchParams
            );
            fileInfo = {
                type: "New",
                driveId: odspResolvedUrl.driveId,
                siteUrl: odspResolvedUrl.siteUrl,
                filePath,
                filename: odspResolvedUrl.fileName,
                createLinkType: createShareLinkParam,
            };
        } else {
            throw new Error(
                "A new or existing file must be specified to create container!"
            );
        }

        if (isCombinedAppAndProtocolSummary(createNewSummary)) {
            const documentAttributes = getDocAttributesFromProtocolSummary(
                createNewSummary.tree[".protocol"]
            );
            if (documentAttributes?.sequenceNumber !== 0) {
                throw new Error(
                    "Seq number in detached ODSP container should be 0"
                );
            }
        }

        const odspLogger = createOdspLogger(logger);

        const fileEntry = {
            resolvedUrl: odspResolvedUrl,
            docId: odspResolvedUrl.hashedDocumentId,
        };
        const cacheAndTracker = createOdspCacheAndTracker(
            this._persistedCache,
            this._nonPersistentCache as any,
            fileEntry,
            odspLogger,
            clientIsSummarizer
        );

        return PerformanceEvent.timedExecAsync(
            odspLogger,
            {
                eventName: "CreateNew",
                isWithSummaryUpload: true,
                createShareLinkParam: createShareLinkParam
                    ? JSON.stringify(createShareLinkParam)
                    : undefined,
                enableShareLinkWithCreate:
                    this._hostPolicy.enableShareLinkWithCreate,
            },
            async (event) => {
                const getStorageToken = toInstrumentedOdspTokenFetcher(
                    odspLogger,
                    resolvedUrlData,
                    this._getStorageToken,
                    true /* throwOnNullToken */
                );
                odspResolvedUrl = isNewFileInfo(fileInfo)
                    ? await createNewFluidFile(
                          getStorageToken,
                          fileInfo,
                          odspLogger,
                          createNewSummary,
                          cacheAndTracker.epochTracker,
                          fileEntry,
                          this._hostPolicy.cacheCreateNewSummary ?? true,
                          !!this._hostPolicy.sessionOptions
                              ?.forceAccessTokenViaAuthorizationHeader,
                          odspResolvedUrl.isClpCompliantApp,
                          this._hostPolicy.enableShareLinkWithCreate
                      )
                    : await createNewContainerOnExistingFile(
                          getStorageToken,
                          fileInfo,
                          odspLogger,
                          createNewSummary,
                          cacheAndTracker.epochTracker,
                          fileEntry,
                          this._hostPolicy.cacheCreateNewSummary ?? true,
                          !!this._hostPolicy.sessionOptions
                              ?.forceAccessTokenViaAuthorizationHeader,
                          odspResolvedUrl.isClpCompliantApp
                      );
                const docService = this.createDocumentServiceCore(
                    odspResolvedUrl,
                    odspLogger,
                    cacheAndTracker as any,
                    clientIsSummarizer
                );
                event.end({
                    docId: odspResolvedUrl.hashedDocumentId,
                });
                return docService;
            }
        );
    }
}

/**
 * Extract the sharing link kind from the resolved URL's query paramerters
 */
function getSharingLinkParams(
    hostPolicy: HostStoragePolicy,
    searchParams: URLSearchParams
): ShareLinkTypes | undefined {
    // extract request parameters for creation of sharing link (if provided) if the feature is enabled
    let createShareLinkParam: ShareLinkTypes | undefined;
    if (hostPolicy.enableShareLinkWithCreate) {
        const createLinkType = searchParams.get("createLinkType");
        if (createLinkType && createLinkType === ShareLinkTypes.csl) {
            createShareLinkParam = ShareLinkTypes.csl;
        }
    }
    return createShareLinkParam;
}
