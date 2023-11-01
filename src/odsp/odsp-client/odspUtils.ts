import { storeLocatorInOdspUrl } from "@fluidframework/odsp-driver";
import {
    IOdspResolvedUrl,
    ShareLinkTypes,
    snapshotKey,
    ICacheEntry,
    TokenFetchOptions,
    OdspErrorType,
    TokenFetcher,
    OdspResourceTokenFetchOptions,
    InstrumentedStorageTokenFetcher,
    tokenFromResponse,
    isTokenFromCache,
} from "@fluidframework/odsp-driver-definitions";
import {
    fetchIncorrectResponse,
    throwOdspNetworkError,
    getSPOAndGraphRequestIdsFromResponse,
} from "@fluidframework/odsp-doclib-utils";
import {
    IResolvedUrl,
    DriverErrorType,
} from "@fluidframework/driver-definitions";
import { Client as MSGraphClient } from "@microsoft/microsoft-graph-client";
import { IDriveInfo } from "./interfaces";
import {
    isOnline,
    OnlineStatus,
    RetryableError,
    NonRetryableError,
    NetworkErrorBasic,
} from "@fluidframework/driver-utils";
import { pkgVersion as driverVersion } from "./packageVersion";
import {
    TelemetryDataTag,
    wrapError,
    PerformanceEvent,
} from "@fluidframework/telemetry-utils";

/**
 * Generates a shareable URL to allow users within a scope to edit a given item in SharePoint.
 * @param msGraphClient Pre-authenticated client object to communicate with the Graph API.
 * @param itemId Unique ID for a resource.
 * @param fileAccessScope Scope of users that will have access to the ODSP fluid file.
 * @param driveInfo Optional parameter containing the site URL and drive ID if the file is not in the
 * current user's personal drive
 */
export async function getShareUrl(
    msGraphClient: MSGraphClient,
    itemId: string,
    fileAccessScope: string,
    driveInfo?: IDriveInfo
): Promise<string> {
    const apiPath: string = driveInfo
        ? `/drives/${driveInfo.driveId}/items/${itemId}/createLink`
        : `/me/drive/items/${itemId}/createLink`;
    return msGraphClient
        .api(apiPath)
        .post({
            type: "edit",
            scope: fileAccessScope,
        })
        .then((rawMessages) => rawMessages.link.webUrl as string);
}

/**
 * Embeds Fluid data store locator data into given ODSP url and returns it.
 * @param shareUrl
 * @param itemId
 * @param driveInfo
 */
export function addLocatorToShareUrl(
    shareUrl: string,
    itemId: string,
    driveInfo: IDriveInfo
): string {
    const shareUrlObject = new URL(shareUrl);
    storeLocatorInOdspUrl(shareUrlObject, {
        siteUrl: driveInfo.siteUrl,
        driveId: driveInfo.driveId,
        itemId,
        dataStorePath: "",
    });
    return shareUrlObject.href;
}

/**
 * Given a Fluid file's metadata information such as its itemId and drive location, return a share link
 * that includes an encoded parameter that contains the necessary information required to locate the
 * file by the ODSP driver
 * @param itemId
 * @param driveInfo
 * @param msGraphClient
 * @param fileAccessScope
 * */
export async function getContainerShareLink(
    itemId: string,
    driveInfo: IDriveInfo,
    msGraphClient: MSGraphClient,
    fileAccessScope = "organization"
): Promise<string> {
    const shareLink = await getShareUrl(
        msGraphClient,
        itemId,
        fileAccessScope,
        driveInfo
    );
    const shareLinkWithLocator = addLocatorToShareUrl(
        shareLink,
        itemId,
        driveInfo
    );

    return shareLinkWithLocator;
}

export interface IFileInfoBase {
    type: "New" | "Existing";
    siteUrl: string;
    driveId: string;
}

export interface INewFileInfo extends IFileInfoBase {
    type: "New";
    filename: string;
    filePath: string;
    /**
     * application can request creation of a share link along with the creation of a new file
     * by passing in an optional param to specify the kind of sharing link
     * (at the time of adding this comment Sept/2021), odsp only supports csl
     * ShareLinkTypes will deprecated in future. Use ISharingLinkKind instead which specifies both
     * share link type and the role type.
     */
    createLinkType?: ShareLinkTypes;
}

export interface IExistingFileInfo extends IFileInfoBase {
    type: "Existing";
    itemId: string;
}

export function isNewFileInfo(
    fileInfo: INewFileInfo | IExistingFileInfo
): fileInfo is INewFileInfo {
    return fileInfo.type === undefined || fileInfo.type === "New";
}

export function getOdspResolvedUrl(
    resolvedUrl: IResolvedUrl
): IOdspResolvedUrl {
    if ((resolvedUrl as IOdspResolvedUrl).odspResolvedUrl !== true) {
        throw 0x1de;
    }
    return resolvedUrl as IOdspResolvedUrl;
}

/**
 * Build request parameters to request for the creation of a sharing link along with the creation of the file
 * through the /snapshot api call.
 * @param shareLinkType - Kind of sharing link requested
 * @returns A string of request parameters that can be concatenated with the base URI
 */
export function buildOdspShareLinkReqParams(
    shareLinkType: ShareLinkTypes | undefined
) {
    if (!shareLinkType) {
        return;
    }
    const scope = (shareLinkType as any).scope;
    if (!scope) {
        return `createLinkType=${shareLinkType}`;
    }
    let shareLinkRequestParams = `createLinkScope=${scope}`;
    const role = (shareLinkType as any).role;
    shareLinkRequestParams = role
        ? `${shareLinkRequestParams}&createLinkRole=${role}`
        : shareLinkRequestParams;
    return shareLinkRequestParams;
}

export function createCacheSnapshotKey(
    odspResolvedUrl: IOdspResolvedUrl
): ICacheEntry {
    const cacheEntry: ICacheEntry = {
        type: snapshotKey,
        key: odspResolvedUrl.fileVersion ?? "",
        file: {
            resolvedUrl: odspResolvedUrl,
            docId: odspResolvedUrl.hashedDocumentId,
        },
    };
    return cacheEntry;
}

export interface TokenFetchOptionsEx extends TokenFetchOptions {
    /** previous error we hit in getWithRetryForTokenRefresh */
    previousError?: any;
}

export const getWithRetryForTokenRefreshRepeat =
    "getWithRetryForTokenRefreshRepeat";

/**
 * This API should be used with pretty much all network calls (fetch, webSocket connection) in order
 * to correctly handle expired tokens. It relies on callback fetching token, and be able to refetch
 * token on failure. Only specific cases get retry call with refresh = true, all other / unknown errors
 * simply propagate to caller
 */
export async function getWithRetryForTokenRefresh<T>(
    get: (options: TokenFetchOptionsEx) => Promise<T>
) {
    return get({ refresh: false }).catch(async (e) => {
        const options: TokenFetchOptionsEx = {
            refresh: true,
            previousError: e,
        };
        switch (e.errorType) {
            // If the error is 401 or 403 refresh the token and try once more.
            case DriverErrorType.authorizationError:
                return get({
                    ...options,
                    claims: e.claims,
                    tenantId: e.tenantId,
                });

            case DriverErrorType.incorrectServerResponse: // some error on the wire, retry once
            case OdspErrorType.fetchTokenError: // If the token was null, then retry once.
                return get(options);

            default:
                // Caller may determine that it wants one retry
                if (e[getWithRetryForTokenRefreshRepeat] === true) {
                    return get(options);
                }
                throw e;
        }
    });
}

/** Parse the given url and return the origin (host name) */
export const getOrigin = (url: string) => new URL(url).origin;

/**
 * @public
 */
export interface IOdspResponse<T> {
    content: T;
    headers: Map<string, string>;
    propsToLog: any;
    duration: number;
}

function headersToMap(headers: Headers) {
    const newHeaders = new Map<string, string>();
    for (const [key, value] of headers.entries()) {
        newHeaders.set(key, value);
    }
    return newHeaders;
}

export async function fetchHelper(
    requestInfo: RequestInfo,
    requestInit: RequestInit | undefined
): Promise<IOdspResponse<Response>> {
    const start = performance.now();

    // Node-fetch and dom have conflicting typing, force them to work by casting for now
    return fetch(requestInfo, requestInit).then(
        async (fetchResponse) => {
            const response = fetchResponse as any as Response;
            // Let's assume we can retry.
            if (!response) {
                throw new NonRetryableError(
                    // pre-0.58 error message: No response from fetch call
                    "No response from ODSP fetch call",
                    DriverErrorType.incorrectServerResponse,
                    { driverVersion }
                );
            }
            if (
                !response.ok ||
                response.status < 200 ||
                response.status >= 300
            ) {
                throwOdspNetworkError(
                    // pre-0.58 error message prefix: odspFetchError
                    `ODSP fetch error [${response.status}]`,
                    response.status,
                    response,
                    await response.text()
                );
            }

            const headers = headersToMap(response.headers);
            return {
                content: response,
                headers,
                propsToLog: getSPOAndGraphRequestIdsFromResponse(headers),
                duration: performance.now() - start,
            };
        },
        (error) => {
            const online = isOnline();

            // The error message may not be suitable to log for privacy reasons, so tag it as such
            const taggedErrorMessage = {
                value: `${error}`, // This uses toString for objects, which often results in `${error.name}: ${error.message}`
                tag: TelemetryDataTag.UserData,
            };
            // After redacting URLs we believe the error message is safe to log
            const urlRegex = /((http|https):\/\/(\S*))/i;
            const redactedErrorText = taggedErrorMessage.value.replace(
                urlRegex,
                "REDACTED_URL"
            );

            // This error is thrown by fetch() when AbortSignal is provided and it gets cancelled
            if (error.name === "AbortError") {
                throw new RetryableError(
                    "Fetch Timeout (AbortError)",
                    OdspErrorType.fetchTimeout,
                    {
                        driverVersion,
                    }
                );
            }
            // TCP/IP timeout
            if (redactedErrorText.includes("ETIMEDOUT")) {
                throw new RetryableError(
                    "Fetch Timeout (ETIMEDOUT)",
                    OdspErrorType.fetchTimeout,
                    {
                        driverVersion,
                    }
                );
            }

            if (online === OnlineStatus.Offline) {
                throw new RetryableError(
                    // pre-0.58 error message prefix: Offline
                    `ODSP fetch failure (Offline): ${redactedErrorText}`,
                    DriverErrorType.offlineError,
                    {
                        driverVersion,
                        rawErrorMessage: taggedErrorMessage,
                    }
                );
            } else {
                // It is perhaps still possible that this is due to being offline, the error does not reveal enough
                // information to conclude.  Could also be DNS errors, malformed fetch request, CSP violation, etc.
                throw new RetryableError(
                    // pre-0.58 error message prefix: Fetch error
                    `ODSP fetch failure: ${redactedErrorText}`,
                    DriverErrorType.fetchFailure,
                    {
                        driverVersion,
                        rawErrorMessage: taggedErrorMessage,
                    }
                );
            }
        }
    );
}

/**
 * A utility function to fetch and parse as JSON with support for retries
 * @param requestInfo - fetch requestInfo, can be a string
 * @param requestInit - fetch requestInit
 */
export async function fetchAndParseAsJSONHelper<T>(
    requestInfo: RequestInfo,
    requestInit: RequestInit | undefined
): Promise<IOdspResponse<T>> {
    const { content, headers, propsToLog, duration } = await fetchHelper(
        requestInfo,
        requestInit
    );
    let text: string | undefined;
    try {
        text = await content.text();
    } catch (e) {
        // JSON.parse() can fail and message would container full request URI, including
        // tokens... It fails for me with "Unexpected end of JSON input" quite often - an attempt to download big file
        // (many ops) almost always ends up with this error - I'd guess 1% of op request end up here... It always
        // succeeds on retry.
        // So do not log error object itself.
        throwOdspNetworkError(
            // pre-0.58 error message: errorWhileParsingFetchResponse
            "Error while parsing fetch response",
            fetchIncorrectResponse,
            content, // response
            text,
            propsToLog
        );
    }

    propsToLog.bodySize = text.length;
    const res = {
        headers,
        content: JSON.parse(text),
        propsToLog,
        duration,
    };
    return res;
}

/**
 * A utility function to fetch and parse as JSON with support for retries
 * @param requestInfo - fetch requestInfo, can be a string
 * @param requestInit - fetch requestInit
 */
export async function fetchArray(
    requestInfo: RequestInfo,
    requestInit: RequestInit | undefined
): Promise<IOdspResponse<ArrayBuffer>> {
    const { content, headers, propsToLog, duration } = await fetchHelper(
        requestInfo,
        requestInit
    );
    let arrayBuffer: ArrayBuffer;
    try {
        arrayBuffer = await content.arrayBuffer();
    } catch (e) {
        // Parsing can fail and message could contain full request URI, including
        // tokens, etc. So do not log error object itself.
        throwOdspNetworkError(
            "Error while parsing fetch response",
            fetchIncorrectResponse,
            content, // response
            undefined, // response text
            propsToLog
        );
    }

    propsToLog.bodySize = arrayBuffer.byteLength;
    return {
        headers,
        content: arrayBuffer,
        propsToLog,
        duration,
    };
}

// 80KB is the max body size that we can put in ump post body for server to be able to accept it.
// Keeping it 78KB to be a little cautious. As per the telemetry 99p is less than 78KB.
export const maxUmpPostBodySize = 79872;

export const createOdspLogger = (logger?: any) => logger;

export function toInstrumentedOdspTokenFetcher(
    logger: any,
    resolvedUrlParts: any,
    tokenFetcher: TokenFetcher<OdspResourceTokenFetchOptions>,
    throwOnNullToken: boolean
): InstrumentedStorageTokenFetcher {
    return async (
        options: TokenFetchOptions,
        name: string,
        alwaysRecordTokenFetchTelemetry: boolean = false
    ) => {
        // Telemetry note: if options.refresh is true, there is a potential perf issue:
        // Host should optimize and provide non-expired tokens on all critical paths.
        // Exceptions: race conditions around expiration, revoked tokens, host that does not care
        // (fluid-fetcher)
        return PerformanceEvent.timedExecAsync(
            logger,
            {
                eventName: `${name}_GetToken`,
                attempts: options.refresh ? 2 : 1,
                hasClaims: !!options.claims,
                hasTenantId: !!options.tenantId,
            },
            async (event) =>
                tokenFetcher({
                    ...options,
                    ...resolvedUrlParts,
                }).then(
                    (tokenResponse) => {
                        const token = tokenFromResponse(tokenResponse);
                        // This event alone generates so many events that is materially impacts cost of telemetry
                        // Thus do not report end event when it comes back quickly.
                        // Note that most of the hosts do not report if result is comming from cache or not,
                        // so we can't rely on that here. But always record if specified explicitly for cases such as
                        // calling trees/latest during load.
                        if (
                            alwaysRecordTokenFetchTelemetry ||
                            event.duration >= 32
                        ) {
                            event.end({
                                fromCache: isTokenFromCache(tokenResponse),
                                isNull: token === null,
                            });
                        }
                        if (token === null && throwOnNullToken) {
                            throw new NonRetryableError(
                                // pre-0.58 error message: Token is null for ${name} call
                                `The Host-provided token fetcher returned null`,
                                OdspErrorType.fetchTokenError,
                                { method: name, driverVersion }
                            );
                        }
                        return token;
                    },
                    (error: any) => {
                        // There is an important but unofficial contract here where token providers can set canRetry: true
                        // to hook into the driver's retry logic (e.g. the retry loop when initiating a connection)
                        const rawCanRetry = error?.canRetry;
                        const tokenError = wrapError(
                            error,
                            (errorMessage) =>
                                new NetworkErrorBasic(
                                    `The Host-provided token fetcher threw an error`,
                                    OdspErrorType.fetchTokenError,
                                    typeof rawCanRetry === "boolean"
                                        ? rawCanRetry
                                        : false /* canRetry */,
                                    {
                                        method: name,
                                        errorMessage,
                                        driverVersion,
                                    }
                                )
                        );
                        throw tokenError;
                    }
                ),
            { cancel: "generic" }
        );
    };
}
