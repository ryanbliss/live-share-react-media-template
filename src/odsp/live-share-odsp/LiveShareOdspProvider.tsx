import { ILiveShareClientOptions, ILiveShareHost } from "@microsoft/live-share";
import {
    FluidContext,
    LiveShareContext,
    DeleteSharedStateAction,
    IAzureContainerResults,
    RegisterSharedSetStateAction,
    SetLocalStateAction,
    UnregisterSharedSetStateAction,
    UpdateSharedStateAction,
} from "@microsoft/live-share-react";
import {
    IFluidContainer,
    LoadableObjectClassRecord,
    IValueChanged,
    SharedMap,
} from "fluid-framework";
import React from "react";
import {
    ILiveShareOdspJoinResults,
    LiveShareOdspClient,
} from "./LiveShareOdspClient";

/**
 * Prop types for {@link LiveShareProvider} component.
 */
export interface ILiveShareProviderProps {
    /**
     * Optional. React children node for the React Context Provider
     */
    children?: React.ReactNode;
    /**
     * Optional. Options for initializing `LiveShareClient`.
     */
    clientOptions?: ILiveShareClientOptions;
    /**
     * Host to initialize `LiveShareClient` with.
     *
     * @remarks
     * If using the `LiveShareClient` class from `@microsoft/teams-js`, you must ensure that you have first called `teamsJs.app.initialize()` before calling `LiveShareClient.create()`.
     */
    host: ILiveShareHost;
    /**
     * The initial object schema to use when {@link joinOnLoad} is true.
     */
    initialObjects?: LoadableObjectClassRecord;
    /**
     * Optional. Flag to determine whether to join Fluid container on load.
     */
    joinOnLoad?: boolean;
}

/**
 * React Context provider component for using Live Share data objects & joining a Live Share session using `LiveShareClient`.
 */
export const LiveShareOdspProvider: React.FC<ILiveShareProviderProps> = (
    props
) => {
    const startedRef = React.useRef(false);
    const clientRef = React.useRef(
        new LiveShareOdspClient(props.host, props.clientOptions)
    );
    const [results, setResults] = React.useState<
        ILiveShareOdspJoinResults | undefined
    >();
    const [joinError, setJoinError] = React.useState<Error | undefined>();

    const stateRegistryCallbacks = useSharedStateRegistry(results);

    /**
     * Join container callback for joining the Live Share session
     */
    const join = React.useCallback(
        async (
            initialObjects?: LoadableObjectClassRecord,
            onInitializeContainer?: (container: IFluidContainer) => void
        ): Promise<ILiveShareOdspJoinResults> => {
            startedRef.current = true;
            const results = await clientRef.current.join(
                initialObjects,
                onInitializeContainer
            );
            setResults(results);
            return results;
        },
        []
    );

    /**
     * Joins the container on load if `props.joinOnLoad` is true
     */
    React.useEffect(() => {
        // This hook should only be called once, so we use a ref to track if it has been called.
        // This is a workaround for the fact that useEffect is called twice on initial render in React V18.
        // We are not doing this here for backwards compatibility. View the README for more information.
        if (results !== undefined || startedRef.current || !props.joinOnLoad)
            return;
        join(props.initialObjects).catch((error) => {
            console.error(error);
            if (error instanceof Error) {
                setJoinError(error);
            } else {
                setJoinError(
                    new Error(
                        "LiveShareProvider: An unknown error occurred while joining container."
                    )
                );
            }
        });
    }, [results, props.joinOnLoad, props.initialObjects, join]);

    return (
        <LiveShareContext.Provider
            value={{
                created: !!results?.created,
                timestampProvider: results?.timestampProvider,
                joined: !!results?.container,
                joinError,
                join,
            }}
        >
            <FluidContext.Provider
                value={{
                    clientRef,
                    container: results?.container,
                    services: results?.services,
                    joinError,
                    getContainer: async () => {
                        throw new Error(
                            "Cannot join new container through getContainer in LiveShareProvider"
                        );
                    },
                    createContainer: async () => {
                        throw new Error(
                            "Cannot create new container through createContainer in LiveShareProvider"
                        );
                    },
                    ...stateRegistryCallbacks,
                }}
            >
                {props.children}
            </FluidContext.Provider>
        </LiveShareContext.Provider>
    );
};

interface ISharedStateRegistryResponse {
    /**
     * Register a set state action callback
     */
    registerSharedSetStateAction: RegisterSharedSetStateAction;
    /**
     * Unregister a set state action callback
     */
    unregisterSharedSetStateAction: UnregisterSharedSetStateAction;
    /**
     * Setter callback to update the shared state
     */
    updateSharedState: UpdateSharedStateAction;
    /**
     * Delete a shared state value
     */
    deleteSharedState: DeleteSharedStateAction;
}

/**
 * Hook used internally to keep track of the SharedSetStateActionMap for each unique key. It sets state values for provided keys and updates components listening to the values.
 *
 * @param results IAzureContainerResults response or undefined
 * @returns ISharedStateRegistryResponse object
 */
const useSharedStateRegistry = (
    results: IAzureContainerResults | undefined
): ISharedStateRegistryResponse => {
    const registeredSharedSetStateActionMapRef = React.useRef<
        Map<string, Map<string, SetLocalStateAction>>
    >(new Map());

    /**
     * @see ISharedStateRegistryResponse.registerSharedSetStateAction
     */
    const registerSharedSetStateAction = React.useCallback(
        (
            uniqueKey: string,
            componentId: string,
            setLocalStateAction: SetLocalStateAction
        ) => {
            let actionsMap =
                registeredSharedSetStateActionMapRef.current.get(uniqueKey);
            if (actionsMap) {
                if (!actionsMap.has(componentId)) {
                    actionsMap.set(componentId, setLocalStateAction);
                }
            } else {
                actionsMap = new Map<string, SetLocalStateAction>();
                actionsMap.set(componentId, setLocalStateAction);
                registeredSharedSetStateActionMapRef.current.set(
                    uniqueKey,
                    actionsMap
                );
            }
            // Set initial values, if known
            const stateMap = results?.container.initialObjects
                .TURBO_STATE_MAP as SharedMap | undefined;
            const initialValue = stateMap?.get(uniqueKey);
            if (initialValue) {
                setLocalStateAction(initialValue);
            }
        },
        [results]
    );

    /**
     * @see ISharedStateRegistryResponse.unregisterSharedSetStateAction
     */
    const unregisterSharedSetStateAction = React.useCallback(
        (uniqueKey: string, componentId: string) => {
            let actionsMap =
                registeredSharedSetStateActionMapRef.current.get(uniqueKey);
            if (actionsMap?.has(componentId)) {
                actionsMap.delete(componentId);
            }
        },
        []
    );

    /**
     * @see ISharedStateRegistryResponse.updateSharedState
     */
    const updateSharedState: UpdateSharedStateAction = React.useCallback(
        (uniqueKey: string, value: any) => {
            if (!results) return;
            const { container } = results;
            const stateMap = container.initialObjects
                .TURBO_STATE_MAP as SharedMap;
            stateMap.set(uniqueKey, value);
        },
        [results]
    );

    /**
     * @see ISharedStateRegistryResponse.deleteSharedState
     */
    const deleteSharedState: DeleteSharedStateAction = React.useCallback(
        (uniqueKey: string) => {
            if (!results) return;
            const { container } = results;
            let actionsMap =
                registeredSharedSetStateActionMapRef.current.get(uniqueKey);
            actionsMap?.clear();
            const stateMap = container.initialObjects
                .TURBO_STATE_MAP as SharedMap;
            stateMap.delete(uniqueKey);
        },
        [results]
    );

    React.useEffect(() => {
        if (!results) return;
        const { container } = results;
        const stateMap = container.initialObjects.TURBO_STATE_MAP as SharedMap;
        const valueChangedListener = (changed: IValueChanged): void => {
            if (registeredSharedSetStateActionMapRef.current.has(changed.key)) {
                const value = stateMap.get(changed.key);
                const actionMap =
                    registeredSharedSetStateActionMapRef.current.get(
                        changed.key
                    );
                actionMap?.forEach((setLocalStateHandler) => {
                    setLocalStateHandler(value);
                });
            }
        };
        stateMap.on("valueChanged", valueChangedListener);
        // Set initial values
        stateMap.forEach((value: any, key: string) => {
            const actionMap =
                registeredSharedSetStateActionMapRef.current.get(key);
            actionMap?.forEach((setLocalStateHandler) => {
                setLocalStateHandler(value);
            });
        });
        return () => {
            stateMap.off("valueChanged", valueChangedListener);
        };
    }, [results]);

    return {
        registerSharedSetStateAction,
        unregisterSharedSetStateAction,
        updateSharedState,
        deleteSharedState,
    };
};
