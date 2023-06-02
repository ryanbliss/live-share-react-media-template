/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { useLiveEvent } from "@microsoft/live-share-react";
import { UNIQUE_KEYS } from "../constants";
import { LivePresenceUser } from "@microsoft/live-share";

/**
 * Hook for sending notifications to display across clients
 */
export const useNotifications = (allUsers: LivePresenceUser[] | undefined) => {
    const { latestEvent, sendEvent } = useLiveEvent<string>(
        UNIQUE_KEYS.notifications
    );

    const userForEvent =
        !!latestEvent &&
        !!allUsers &&
        allUsers.length > 0 &&
        allUsers.find((user) => !!user.getConnection(latestEvent.clientId));

    return {
        notificationToDisplay:
            !!latestEvent && !!userForEvent
                ? `${
                      userForEvent.isLocalUser
                          ? "You"
                          : userForEvent.displayName
                  } ${latestEvent.value}`
                : undefined,
        sendNotification: sendEvent,
    };
};
