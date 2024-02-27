/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { IMember, IServiceAudience } from "@fluidframework/fluid-static";
import { ITelemetryBaseLogger } from "@fluidframework/core-interfaces";
import { ITokenProvider } from "@fluidframework/azure-client";
import { OdspMember } from "@fluid-experimental/odsp-client";

/**
 * OdspConnectionConfig defines the necessary properties that will be applied to all containers
 * created by an OdspClient instance. This includes callbacks for the authentication tokens
 * required for ODSP.
 */
export interface OdspConnectionConfig {
    /**
     * Site url representing ODSP resource location
     */
    siteUrl: string;

    /**
     * Instance that provides AAD endpoint tokens for Push and SharePoint
     */
    tokenProvider: ITokenProvider;

    /**
     * RaaS Drive Id of the tenant where Fluid containers are created
     */
    driveId: string;
}

export interface OdspClientProps {
    /**
     * Configuration for establishing a connection with the ODSP Fluid Service (Push).
     */
    readonly connection: OdspConnectionConfig;

    /**
     * Optional. A logger instance to receive diagnostic messages.
     */
    readonly logger?: ITelemetryBaseLogger;

    /**
     * Base interface for providing configurations to control experimental features. If unsure, leave this undefined.
     */
    // readonly configProvider?: IConfigProviderBase;
}

export const tokenMap: Map<string, string> = new Map();

/**
 * OdspContainerServices is returned by the OdspClient alongside a FluidContainer. It holds the
 * functionality specifically tied to the ODSP service, and how the data stored in the
 * FluidContainer is persisted in the backend and consumed by users. Any functionality regarding
 * how the data is handled within the FluidContainer itself, i.e. which data objects or DDSes to
 * use, will not be included here but rather on the FluidContainer class itself.
 */
export interface OdspContainerServicesExt {
    /**
     * Provides an object that can be used to get the users that are present in this Fluid session and
     * listeners for when the roster has any changes from users joining/leaving the session
     */
    audience: IOdspAudienceExt;
}

export interface OdspUser {
    /**
     * The user's name
     */
    name: string;

    email: string;

    /**
     * The object ID or object Identifier. It is a unique identifier assigned to each user, group, or other entity within AAD or another Microsoft 365 service.
     * It is a GUID that uniquely identifies the object. When making Microsoft Graph API calls, you might need to reference or manipulate objects within the directory, and the `oid` is used to identify these objects.
     */
    oid: string;
}

/**
 * Since ODSP provides user names and email for all of its members, we extend the
 * {@link @fluidframework/protocol-definitions#IMember} interface to include this service-specific value.
 * It will be returned for all audience members connected.
 * @beta
 */
export interface OdspMemberExt extends OdspMember {
    /**
     * The user's name
     */
    userName: string;
}

/**
 * Audience object for ODSP containers
 * @beta
 */
export type IOdspAudienceExt = IServiceAudience<OdspMemberExt>;
