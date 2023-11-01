/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import {
    PublicClientApplication,
    AuthenticationResult,
    InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { tokenMap } from "../odsp-client";

const msalConfig = {
    auth: {
        clientId: "19abc360-c059-48d8-854e-cfeef9a3c5b8",
        authority: "https://login.microsoftonline.com/common/",
    },
};

const graphScopes = ["FileStorageContainer.Selected"];

const sharePointScopes = [
    "https://M365x82694150.sharepoint.com/Container.Selected",
    "https://M365x82694150.sharepoint.com/AllSites.Write",
    // "https://M365x82694150-my.sharepoint.com/personal/admin_m365x82694150_onmicrosoft_com/Container.Selected",
    // "https://M365x82694150-my.sharepoint.com/personal/admin_m365x82694150_onmicrosoft_com/AllSites.Write",
    // "https://M365x82694150-my.sharepoint.com/Container.Selected",
    // "https://M365x82694150-my.sharepoint.com/AllSites.Write",
];

const pushScopes = [
    "offline_access",
    "https://pushchannel.1drv.ms/PushChannel.ReadWrite.All",
];

const msalInstance = new PublicClientApplication(msalConfig);

export async function getTokens(): Promise<{
    graphToken: string;
    sharePointToken: string;
    pushToken: string;
    userName: string;
    siteUrl: string;
}> {
    const response = await msalInstance.loginPopup({ scopes: graphScopes });

    msalInstance.setActiveAccount(response.account);
    const username = response.account?.username as string;
    const startIndex = username.indexOf("@") + 1;
    const endIndex = username.indexOf(".");
    const tenantName = username.substring(startIndex, endIndex);
    const siteUrl = `https://${tenantName}.sharepoint.com`;
    // const siteUrl = `https://${tenantName}-my.sharepoint.com/personal/admin_m365x82694150_onmicrosoft_com`;

    try {
        // Attempt to acquire SharePoint token silently
        const sharePointRequest = {
            scopes: sharePointScopes,
        };
        const sharePointTokenResult: AuthenticationResult =
            await msalInstance.acquireTokenSilent(sharePointRequest);

        // Attempt to acquire other token silently
        const otherRequest = {
            scopes: pushScopes,
        };
        const pushTokenResult: AuthenticationResult =
            await msalInstance.acquireTokenSilent(otherRequest);

        tokenMap.set("graphToken", response.accessToken);
        tokenMap.set("sharePointToken", sharePointTokenResult.accessToken);
        tokenMap.set("pushToken", pushTokenResult.accessToken);
        tokenMap.set("userName", username);
        tokenMap.set("siteUrl", siteUrl);

        // Return both tokens
        return {
            graphToken: response.accessToken,
            sharePointToken: sharePointTokenResult.accessToken,
            pushToken: pushTokenResult.accessToken,
            userName: response.account?.username as string,
            siteUrl: siteUrl,
        };
    } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
            // If silent token acquisition fails, fall back to interactive flow
            const sharePointRequest = {
                scopes: sharePointScopes,
            };
            const sharePointTokenResult: AuthenticationResult =
                await msalInstance.acquireTokenPopup(sharePointRequest);

            const otherRequest = {
                scopes: pushScopes,
            };
            const pushTokenResult: AuthenticationResult =
                await msalInstance.acquireTokenPopup(otherRequest);

            tokenMap.set("graphToken", response.accessToken);
            tokenMap.set("sharePointToken", sharePointTokenResult.accessToken);
            tokenMap.set("pushToken", pushTokenResult.accessToken);
            tokenMap.set("userName", username);
            tokenMap.set("siteUrl", siteUrl);

            // Return both tokens
            return {
                graphToken: response.accessToken,
                sharePointToken: sharePointTokenResult.accessToken,
                pushToken: pushTokenResult.accessToken,
                userName: response.account?.username as string,
                siteUrl: siteUrl,
            };
        } else {
            // Handle any other error
            console.error(error);
            throw error;
        }
    }
}
