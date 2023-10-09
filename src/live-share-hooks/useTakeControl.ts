import { useCallback, useMemo } from "react";
import { LivePresenceUser } from "@microsoft/live-share";
import { IUserData } from "./usePresence";
import {
    SendLiveEventAction,
    useLiveShareContext,
    useSharedMap,
} from "@microsoft/live-share-react";
import { UNIQUE_KEYS } from "../constants";

export const useTakeControl = (
    localUser: LivePresenceUser<IUserData> | undefined,
    localUserIsEligiblePresenter: boolean,
    users: LivePresenceUser<IUserData>[],
    sendNotification: SendLiveEventAction<string>
) => {
    const { sharedMap: takeControlMap, map: history } = useSharedMap<number>(
        UNIQUE_KEYS.takeControl
    );
    const { timestampProvider } = useLiveShareContext();

    // Computed presentingUser object based on most recent online user to take control
    const presentingUser = useMemo(() => {
        const mappedUsers = users.map((user) => {
            return {
                userId: user.userId,
                state: user.state,
                data: user.data,
                lastInControlTimestamp: user.userId
                    ? history.get(user.userId)
                    : 0,
            };
        });
        mappedUsers.sort((a, b) => {
            // Sort by joined timestamp in descending
            if (a.lastInControlTimestamp === b.lastInControlTimestamp) {
                return (
                    (a.data?.joinedTimestamp ?? 0) -
                    (b.data?.joinedTimestamp ?? 0)
                );
            }
            // Sort by last in control time in ascending
            return (
                (b.lastInControlTimestamp ?? 0) -
                (a.lastInControlTimestamp ?? 0)
            );
        });
        return mappedUsers[0];
    }, [history, users]);

    // Local user is the presenter
    const localUserIsPresenting = useMemo(() => {
        if (!presentingUser || !localUser) {
            return false;
        }
        return localUser.userId === presentingUser.userId;
    }, [localUser, presentingUser]);

    // Set the local user ID
    const takeControl = useCallback(() => {
        if (!!localUser?.userId && localUserIsEligiblePresenter) {
            takeControlMap?.set(
                localUser?.userId,
                timestampProvider?.getTimestamp()
            );
            sendNotification?.("took control");
        }
    }, [
        takeControlMap,
        localUser,
        localUserIsEligiblePresenter,
        timestampProvider,
        sendNotification,
    ]);

    return {
        takeControlStarted: !!takeControlMap,
        presentingUser,
        localUserIsPresenting,
        takeControl,
    };
};
