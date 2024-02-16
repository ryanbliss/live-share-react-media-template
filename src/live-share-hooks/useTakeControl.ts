import { useCallback, useEffect, useRef } from "react";
import {
    FollowModePresenceUser,
    FollowModeType,
    IFollowModeState,
    LiveDataObjectInitializeState,
    LiveFollowMode,
} from "@microsoft/live-share";
import {
    useFluidObjectsContext,
    useLiveFollowMode,
} from "@microsoft/live-share-react";
import {
    ACCEPT_PLAYBACK_CHANGES_FROM,
    AppConfiguration,
    UNIQUE_KEYS,
} from "../constants";
import { DisplayNotificationCallback } from "./useNotifications";
import { MediaItem, searchList } from "../utils/media-list";

export interface IFollowModeData {
    mediaId: string | undefined;
    paused: boolean;
    changed?: {
        timestamp: number;
        mediaPosition: number;
    };
}

export type FollowModeState = IFollowModeState<IFollowModeData> | undefined;
export type UpdateFollowStateCallback = (
    state: IFollowModeData
) => Promise<void>;
export type FollowUserStateCallback = (userId: string) => Promise<void>;
export type FollowModeUser = FollowModePresenceUser<IFollowModeData>;
export type TLiveFollowMode = LiveFollowMode<IFollowModeData> | undefined;

export const useTakeControl = (
    threadId: string,
    isShareInitiator: boolean,
    displayNotification: DisplayNotificationCallback,
    initialMediaItem: MediaItem | undefined
) => {
    const { clientRef } = useFluidObjectsContext();
    const {
        state,
        localUser,
        otherUsers,
        followUser,
        startPresenting,
        stopPresenting,
        stopFollowing,
        endSuspension,
        beginSuspension,
        update,
        liveFollowMode,
    } = useLiveFollowMode<IFollowModeData>(
        `${threadId}/${UNIQUE_KEYS.takeControl}`,
        {
            mediaId: initialMediaItem?.id,
            paused: true,
        },
        ACCEPT_PLAYBACK_CHANGES_FROM
    );

    // Local user is the presenter
    const localUserIsPresenting = state?.type
        ? [
              FollowModeType.activePresenter,
              FollowModeType.activeFollowers,
              FollowModeType.local,
          ].includes(state.type)
        : false;
    const localUserIsEligiblePresenter = localUser
        ? localUser.roles.filter((role) =>
              ACCEPT_PLAYBACK_CHANGES_FROM.includes(role)
          ).length > 0
        : false;

    // Set the local user ID
    const takeControl = useCallback(() => {
        if (!liveFollowMode) return;
        if (!localUser) return;
        if (!state) return;
        if (state.type === FollowModeType.activePresenter) return;
        if (!localUserIsEligiblePresenter) return;
        startPresenting().catch((err) => {
            console.error(err);
        });
    }, [
        localUserIsEligiblePresenter,
        localUser,
        liveFollowMode,
        state,
        startPresenting,
        displayNotification,
    ]);

    // Start presenting if nobody is in control and local user isShareInitiator
    useEffect(() => {
        if (
            liveFollowMode?.initializeState !==
            LiveDataObjectInitializeState.succeeded
        )
            return;
        if (!state) return;
        if (!isShareInitiator) return;
        if (state.type !== FollowModeType.local) return;
        startPresenting();
    }, [
        isShareInitiator,
        state,
        liveFollowMode?.initializeState,
        startPresenting,
    ]);

    // Set canSendBackgroundUpdates when localUserIsPresenting changes
    useEffect(() => {
        if (!AppConfiguration.isFullyLargeMeetingOptimized) return;
        // We will allow both the presenter and isShareInitiator to send background updates.
        // These will be the same as long as nobody has taken control since presenting started.
        // We want at least one client to always have this value == true so that `connect` events are not missed.
        // Since `liveFollowMode` starts out with no presenter, this will ensure at least one client is always in a `connect` state.
        // This is safer even it this may mean two clients are responding to `connect` events / sending background events.
        // When not in Teams, this is currently set to always be true, since `isShareInitiator` is hardcoded to true.
        clientRef.current.canSendBackgroundUpdates =
            localUserIsPresenting || isShareInitiator;
    }, [clientRef, localUserIsPresenting, isShareInitiator]);

    // Display take control notification
    const hasSkippedFirstRef = useRef(false);
    useEffect(() => {
        if (!liveFollowMode) return;
        if (!state) return;
        const user = liveFollowMode.getUser(state.followingUserId);
        if (!user) return;
        const userConnections = user.getConnections();
        if (userConnections.length === 0) return;
        if (
            ![
                FollowModeType.activePresenter,
                FollowModeType.followPresenter,
            ].includes(state.type)
        )
            return;
        // Skip the first notification when the app first loads
        if (!hasSkippedFirstRef.current) {
            hasSkippedFirstRef.current = true;
            return;
        }
        displayNotification(
            liveFollowMode,
            user.isLocalUser ? "are in control" : "is in control",
            userConnections[0].clientId,
            user.isLocalUser
        );
    }, [state?.type, liveFollowMode]);

    const selectedMediaItem = searchList.find(
        (item) => item.id === state?.value?.mediaId
    );

    return {
        localUser,
        otherUsers,
        takeControlStarted: !!liveFollowMode,
        localUserIsEligiblePresenter,
        localUserIsPresenting,
        selectedMediaItem,
        state,
        liveFollowMode,
        updateFollowState: update,
        followUser,
        takeControl,
        stopFollowing,
        stopPresenting,
        beginSuspension,
        endSuspension,
    };
};
