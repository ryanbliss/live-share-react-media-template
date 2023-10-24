import {
    ConnectionState,
    ContainerSchema,
    FluidContainer,
    IFluidContainer,
    LoadableObjectClassRecord,
} from "fluid-framework";
import {
    OdspClient,
    OdspContainerServices,
    OdspCreateContainerConfig,
    OdspGetContainerConfig,
    getOdspDriver,
} from "../odsp-client";
import { v4 as uuid } from "uuid";
import {
    ILiveShareClientOptions,
    ILiveShareHost,
    ILiveShareJoinResults,
    LiveShareRuntime,
    LocalTimestampProvider,
    TestLiveShareHost,
    getLiveShareContainerSchemaProxy,
} from "@microsoft/live-share";
import { FluidTurboClient } from "@microsoft/live-share-turbo";
import { AzureContainerServices } from "@fluidframework/azure-client";

const documentId = uuid();

export class LiveShareOdspClient extends FluidTurboClient {
    private _host: ILiveShareHost;
    private readonly _runtime: LiveShareRuntime;
    private readonly _options: ILiveShareClientOptions;
    private _results: ILiveShareJoinResults | undefined;
    /**
     * Creates a new `LiveShareClient` instance.
     * @param host Host for the current Live Share session.
     * @param options Optional. Configuration options for the client.
     */
    constructor(host: ILiveShareHost, options?: ILiveShareClientOptions) {
        super();
        // Validate host passed in
        if (!host) {
            throw new Error(
                `LiveShareClient: prop \`host\` is \`${host}\` when it is expected to be a non-optional value of type \`ILiveShareHost\`. Please ensure \`host\` is defined before initializing \`LiveShareClient\`.`
            );
        }
        if (typeof host.getFluidTenantInfo != "function") {
            throw new Error(
                `LiveShareClient: \`host.getFluidTenantInfo\` is of type \`${typeof host.getFluidTenantInfo}\` when it is expected to be a type of \`function\`. For more information, review the \`ILiveShareHost\` interface.`
            );
        }
        this._host = host;
        // Save options
        this._options = {
            ...options,
            timestampProvider: getIsTestClient(host, options)
                ? new LocalTimestampProvider()
                : options?.timestampProvider,
        };
        this._runtime = new LiveShareRuntime(this._host, this._options, true);
    }

    /**
     * If true the client is configured to use a local test server.
     */
    public get isTesting(): boolean {
        return getIsTestClient(this._host, this._options);
    }

    /**
     * Number of times the client should attempt to get the ID of the container to join for the
     * current context.
     */
    public maxContainerLookupTries = 3;

    /**
     * Get the Fluid join container results
     */
    public override get results(): ILiveShareJoinResults | undefined {
        return this._results;
    }

    /**
     * Setting for whether `LiveDataObject` instances using `LiveObjectSynchronizer` can send background updates.
     * Default value is `true`.
     *
     * @remarks
     * This is useful for scenarios where there are a large number of participants in a session, since service performance degrades as more socket connections are opened.
     * Intended for use when a small number of users are intended to be "in control", such as the `LiveFollowMode` class's `startPresenting()` feature.
     * There should always be at least one user in the session that has `canSendBackgroundUpdates` set to true.
     * Set to true when the user is eligible to send background updates (e.g., "in control"), or false when that user is not in control.
     * This setting will not prevent the local user from explicitly changing the state of objects using `LiveObjectSynchronizer`, such as `.set()` in `LiveState`.
     * Impacts background updates of `LiveState`, `LivePresence`, `LiveTimer`, and `LiveFollowMode`.
     */
    public get canSendBackgroundUpdates(): boolean {
        return this._runtime.canSendBackgroundUpdates;
    }

    public set canSendBackgroundUpdates(value: boolean) {
        this._runtime.canSendBackgroundUpdates = value;
    }

    async join(
        initialObjects?: LoadableObjectClassRecord,
        onContainerFirstCreated?: (container: IFluidContainer) => void
    ): Promise<ILiveShareJoinResults> {
        // Apply runtime to ContainerSchema
        const containerSchema = getLiveShareContainerSchemaProxy(
            this.getContainerSchema(initialObjects),
            this._runtime
        );
        console.log(
            "LiveShareOdspClient::joinContainer: initiating the driver"
        );
        const odspDriver = await getOdspDriver();
        console.log(
            "LiveShareOdspClient::joinContainer: initial driver",
            odspDriver
        );

        const { containerId, isNew } = this.getContainerId();

        let container: FluidContainer;
        let services: OdspContainerServices;

        if (isNew) {
            console.log(
                "LiveShareOdspClient::joinContainer: creating the container"
            );
            const containerConfig: OdspCreateContainerConfig = {
                siteUrl: odspDriver.siteUrl,
                driveId: odspDriver.driveId,
                folderName: odspDriver.directory,
                fileName: documentId,
            };

            console.log(
                "LiveShareOdspClient::joinContainer: container config",
                containerConfig
            );

            const { fluidContainer, containerServices } =
                await OdspClient.createContainer(
                    containerConfig,
                    containerSchema
                );
            onContainerFirstCreated?.(fluidContainer);
            container = fluidContainer;
            services = containerServices;

            const sharingLink = await containerServices.generateLink();
            const itemId = containerPath(sharingLink);
            localStorage.setItem(itemId, sharingLink);
            console.log(
                "LiveShareOdspClient::joinContainer: container created"
            );
            location.hash = itemId;
        } else {
            const containerConfig: OdspGetContainerConfig = {
                fileUrl: containerId, //pass file url
            };

            const { fluidContainer, containerServices } =
                await OdspClient.getContainer(containerConfig, containerSchema);

            container = fluidContainer;
            services = containerServices;
        }

        if (container.connectionState !== ConnectionState.Connected) {
            await new Promise<void>((resolve) => {
                container.once("connected", () => {
                    resolve();
                });
            });
        }

        const results = {
            container,
            // TODO: fix
            services: services as AzureContainerServices,
            timestampProvider: this._runtime.timestampProvider,
            created: isNew,
        };
        this._results = results;
        return results;
    }

    private getContainerId(): { containerId: string; isNew: boolean } {
        let isNew = false;
        console.log(
            "LiveShareOdspClient::getContainerId: hash: ",
            location.hash
        );
        if (location.hash.length === 0) {
            isNew = true;
        }
        const hash = location.hash;
        const itemId = hash.charAt(0) === "#" ? hash.substring(1) : hash;
        const containerId = localStorage.getItem(itemId)!;
        return { containerId, isNew };
    }
}

function containerPath(url: string): string {
    const itemIdPattern = /itemId=([^&]+)/; // regular expression to match the itemId parameter value
    let itemId;

    const match = url.match(itemIdPattern); // get the match object for the itemId parameter value
    if (match) {
        itemId = match[1]; // extract the itemId parameter value from the match object
        console.log(itemId); // output: "itemidQ"
    } else {
        console.log("itemId parameter not found in the URL");
        itemId = "";
    }
    return itemId;
}

function getIsTestClient(
    host: ILiveShareHost,
    options?: ILiveShareClientOptions
) {
    return (
        options?.connection?.type == "local" ||
        host instanceof TestLiveShareHost
    );
}