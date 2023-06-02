/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { useMemo } from "react";
import {
    useLivePresence,
    useLiveShareContext,
} from "@microsoft/live-share-react";
import { ACCEPT_PLAYBACK_CHANGES_FROM, UNIQUE_KEYS } from "../constants";

export interface IUserData {
    joinedTimestamp: number;
}

/**
 * Hook for tracking users
 */
export const usePresence = () => {
    const { timestampProvider } = useLiveShareContext();
    const { allUsers, localUser, livePresence } = useLivePresence<IUserData>(
        UNIQUE_KEYS.presence,
        // Get initial value callback
        () => ({
            joinedTimestamp: timestampProvider?.getTimestamp() ?? 0,
        })
    );

    // Local user is an eligible presenter
    const localUserIsEligiblePresenter = useMemo(() => {
        if (ACCEPT_PLAYBACK_CHANGES_FROM.length === 0) {
            return true;
        }
        if (!livePresence || !localUser) {
            return false;
        }
        return (
            localUser.roles.filter((role) =>
                ACCEPT_PLAYBACK_CHANGES_FROM.includes(role)
            ).length > 0
        );
    }, [livePresence, localUser]);

    return {
        presenceStarted: !!livePresence,
        localUser,
        allUsers,
        localUserIsEligiblePresenter,
    };
};
