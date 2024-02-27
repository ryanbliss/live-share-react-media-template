/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */
import { v4 as uuid } from "uuid";
import {
    AttachState,
    IContainer,
    IFluidModuleWithDetails,
} from "@fluidframework/container-definitions";
import { FluidObject, IRequest } from "@fluidframework/core-interfaces";
import { assert } from "@fluidframework/core-utils";
import { Loader } from "@fluidframework/container-loader";
import { IDocumentServiceFactory } from "@fluidframework/driver-definitions";
import {
    type ContainerSchema,
    createDOProviderContainerRuntimeFactory,
    IFluidContainer,
    createFluidContainer,
    IRootDataObject,
    createServiceAudience,
} from "@fluidframework/fluid-static";
import {
    OdspDocumentServiceFactory,
    OdspDriverUrlResolver,
    createOdspCreateContainerRequest,
    createOdspUrl,
    isOdspResolvedUrl,
} from "@fluidframework/odsp-driver";
import type {
    IOdspResolvedUrl,
    OdspResourceTokenFetchOptions,
    TokenResponse,
} from "@fluidframework/odsp-driver-definitions";
import { IClient } from "@fluidframework/protocol-definitions";
import {
    OdspClientProps,
    OdspContainerServices,
    OdspConnectionConfig,
    IOdspTokenProvider,
    OdspMember,
} from "@fluid-experimental/odsp-client";
import { OdspContainerServicesExt, OdspMemberExt } from "./interfaces";

/**
 * Since ODSP provides user names, email and oids for all of its members, we extend the
 * {@link @fluidframework/fluid-static#IMember} interface to include this service-specific value.
 * @internal
 */
interface OdspUser {
    /**
     * The user's email address
     */
    email: string;
    /**
     * The user's name
     */
    name: string;
    /**
     * The object ID (oid). It is a unique identifier assigned to each user, group, or other entity within AAD or another Microsoft 365 service. It is a GUID that uniquely identifies the object. When making Microsoft Graph API calls, you might need to reference or manipulate objects within the directory, and the `oid` is used to identify these objects.
     */
    oid: string;
}

function createOdspAudienceMember(audienceMember: IClient): OdspMemberExt {
    const user = audienceMember.user as unknown as OdspUser;
    assert(
        user.name !== undefined ||
            user.email !== undefined ||
            user.oid !== undefined,
        0x836 /* Provided user was not an "OdspUser". */
    );

    return {
        userId: user.oid,
        name: user.name,
        userName: user.name,
        email: user.email,
        connections: [],
    };
}

async function getStorageToken(
    options: OdspResourceTokenFetchOptions,
    tokenProvider: IOdspTokenProvider
): Promise<TokenResponse> {
    const tokenResponse: TokenResponse = await tokenProvider.fetchStorageToken(
        options.siteUrl,
        options.refresh
    );
    return tokenResponse;
}

async function getWebsocketToken(
    options: OdspResourceTokenFetchOptions,
    tokenProvider: IOdspTokenProvider
): Promise<TokenResponse> {
    const tokenResponse: TokenResponse =
        await tokenProvider.fetchWebsocketToken(
            options.siteUrl,
            options.refresh
        );
    return tokenResponse;
}

/**
 * OdspClient provides the ability to have a Fluid object backed by the ODSP service within the context of Microsoft 365 (M365) tenants.
 * @sealed
 * @beta
 */
export class OdspClient {
    private readonly documentServiceFactory: IDocumentServiceFactory;
    private readonly urlResolver: OdspDriverUrlResolver;

    public constructor(private readonly properties: OdspClientProps) {
        this.documentServiceFactory = new OdspDocumentServiceFactory(
            async (options) =>
                getStorageToken(
                    options,
                    this.properties.connection.tokenProvider
                ),
            async (options) =>
                getWebsocketToken(
                    options,
                    this.properties.connection.tokenProvider
                )
        );

        this.urlResolver = new OdspDriverUrlResolver();
    }

    public async createContainer(containerSchema: ContainerSchema): Promise<{
        container: IFluidContainer;
        services: OdspContainerServicesExt;
    }> {
        const loader = this.createLoader(containerSchema);

        const container = await loader.createDetachedContainer({
            package: "no-dynamic-package",
            config: {},
        });

        const fluidContainer = await this._createNewFluidContainer(
            container,
            this.properties.connection
        );

        const services = await this.getContainerServices(container);

        return { container: fluidContainer, services };
    }

    public async getContainer(
        id: string,
        containerSchema: ContainerSchema
    ): Promise<{
        container: IFluidContainer;
        services: OdspContainerServicesExt;
    }> {
        const loader = this.createLoader(containerSchema);
        const url = createOdspUrl({
            siteUrl: this.properties.connection.siteUrl,
            driveId: this.properties.connection.driveId,
            itemId: id,
            dataStorePath: "",
        });
        const container = await loader.resolve({ url });

        const fluidContainer = createFluidContainer({
            container,
            rootDataObject: await this.getContainerEntryPoint(container),
        });
        const services = await this.getContainerServices(container);
        return { container: fluidContainer, services };
    }

    public async createContainerForExistingFile(
        id: string,
        containerSchema: ContainerSchema
    ): Promise<{
        container: IFluidContainer;
        services: OdspContainerServicesExt;
    }> {
        const url = createOdspUrl({
            siteUrl: this.properties.connection.siteUrl,
            driveId: this.properties.connection.driveId,
            itemId: id,
            dataStorePath: "",
        });
        const loader = this.createLoader(containerSchema);

        // We're not actually using the code proposal (our code loader always loads the same module regardless of the
        // proposal), but the Container will only give us a NullRuntime if there's no proposal.  So we'll use a fake
        // proposal.
        let container = await loader.createDetachedContainer({
            package: "no-dynamic-package",
            config: {},
        });
        const rootDataObject = await this.getContainerEntryPoint(container);
        /**
         * See {@link FluidContainer.attach}
         */
        const attach = async (): Promise<string> => {
            let request: IRequest = {
                url,
                headers: {
                    // [LoaderHeader.loadMode]: {
                    //     /*
                    //      * Connection to delta stream is made only when Container.connect() call is made.
                    //      * Op fetching from storage is performed and ops are applied as they come in.
                    //      * This is useful option if connection to delta stream is expensive and thus it's beneficial to move it
                    //      * out from critical boot sequence, but it's beneficial to allow catch up to happen as fast as possible.
                    //      */
                    //     deltaConnection: "none",
                    // },
                },
            };

            if (container.attachState !== AttachState.Detached) {
                throw new Error(
                    "Cannot attach container. Container is not in detached state"
                );
            }
            await container.attach(request);
            const resolvedUrl = container.resolvedUrl;
            if (resolvedUrl === undefined || !isOdspResolvedUrl(resolvedUrl)) {
                throw new Error(
                    "Resolved Url not available on attached container"
                );
            }
            if (container.resolvedUrl === undefined) {
                throw new Error(
                    "Resolved Url not available on attached container"
                );
            }
            return resolvedUrl.itemId;
        };
        const fluidContainer = createFluidContainer({
            container,
            rootDataObject,
        });
        fluidContainer.attach = attach;
        const services = await this.getContainerServices(container);
        return {
            container: fluidContainer,
            services,
        };
    }

    private createLoader(schema: ContainerSchema): Loader {
        const runtimeFactory = createDOProviderContainerRuntimeFactory({
            schema,
        });
        const load = async (): Promise<IFluidModuleWithDetails> => {
            return {
                module: { fluidExport: runtimeFactory },
                details: { package: "no-dynamic-package", config: {} },
            };
        };

        const codeLoader = { load };
        const client: IClient = {
            details: {
                capabilities: { interactive: true },
            },
            permission: [],
            scopes: [],
            user: { id: "" },
            mode: "write",
        };

        return new Loader({
            urlResolver: this.urlResolver,
            documentServiceFactory: this.documentServiceFactory,
            codeLoader,
            logger: this.properties.logger,
            options: { client },
        });
    }

    private async _createNewFluidContainer(
        container: IContainer,
        connection: OdspConnectionConfig
    ): Promise<IFluidContainer> {
        const rootDataObject = await this.getContainerEntryPoint(container);

        /**
         * See {@link FluidContainer.attach}
         */
        const attach = async (odspProps?: any): Promise<string> => {
            const createNewRequest: IRequest = createOdspCreateContainerRequest(
                connection.siteUrl,
                connection.driveId,
                odspProps?.filePath ?? "",
                odspProps?.fileName ?? uuid()
            );
            if (container.attachState !== AttachState.Detached) {
                throw new Error(
                    "Cannot attach container. Container is not in detached state"
                );
            }
            await container.attach(createNewRequest);

            const resolvedUrl = container.resolvedUrl;

            if (resolvedUrl === undefined || !isOdspResolvedUrl(resolvedUrl)) {
                throw new Error(
                    "Resolved Url not available on attached container"
                );
            }

            /**
             * A unique identifier for the file within the provided RaaS drive ID. When you attach a container,
             * a new `itemId` is created in the user's drive, which developers can use for various operations
             * like updating, renaming, moving the Fluid file, changing permissions, and more. `itemId` is used to load the container.
             */
            return resolvedUrl.itemId;
        };
        const fluidContainer = createFluidContainer({
            container,
            rootDataObject,
        });
        fluidContainer.attach = attach;
        return fluidContainer;
    }

    private async getContainerServices(
        container: IContainer
    ): Promise<OdspContainerServicesExt> {
        return {
            audience: createServiceAudience({
                container,
                createServiceMember: createOdspAudienceMember,
            }),
        };
    }

    private async getContainerEntryPoint(
        container: IContainer
    ): Promise<IRootDataObject> {
        const rootDataObject: FluidObject<IRootDataObject> =
            await container.getEntryPoint();
        assert(
            rootDataObject.IRootDataObject !== undefined,
            "Invalid IRootDataObject reference" /* entryPoint must be of type IRootDataObject */
        );
        return rootDataObject.IRootDataObject;
    }
}
