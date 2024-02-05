/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { getTokens } from "../msal/tokens";
import { OdspConnectionConfig } from "./interfaces";
import { OdspClient } from "./OdspClient";
import { OdspDriver } from "./OdspDriver";

const initDriver = async (spoToken: string): Promise<OdspDriver> => {
    console.log("Driver init------");

    // const { graphToken, sharePointToken, pushToken, userName, siteUrl } =
    //     await getTokens();
    // console.log(
    //     "InitiateDriver::initDriver: tokens-------------------" + graphToken,
    //     sharePointToken,
    //     pushToken,
    //     userName,
    //     siteUrl
    // );

    // TODO: get upn from decoded token
    const userName = "blah@placeholder.com";

    const driver: OdspDriver = await OdspDriver.createFromEnv({
        username: userName,
        // directory: "Sonali-Brainstorm-1",
        supportsBrowserAuth: true,
        odspEndpointName: "odsp",
    });
    console.log("InitiateDriver::initDriver: Driver------", driver);
    const connectionConfig: OdspConnectionConfig = {
        getSharePointToken: driver.getStorageToken as any,
        getPushServiceToken: driver.getPushToken as any,
        getGraphToken: driver.getGraphToken as any,
        getMicrosoftGraphToken: spoToken,
    };

    OdspClient.init(connectionConfig, driver.siteUrl);
    return driver;
};

export const getOdspDriver = async (spoToken: string): Promise<OdspDriver> => {
    const odspDriver = await initDriver(spoToken);
    console.log("InitiateDriver:: getOdspDriver: INITIAL DRIVER", odspDriver);
    return odspDriver;
};
