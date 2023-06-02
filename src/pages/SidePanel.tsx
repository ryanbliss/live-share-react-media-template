/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { FC, useCallback, useEffect, useRef } from "react";
import { useTeamsContext } from "../teams-js-hooks/useTeamsContext";
import { useNavigate } from "react-router-dom";
import { MediaItem, mediaList, searchList } from "../utils/media-list";
import { ListWrapper, LiveSharePage } from "../components";
import * as liveShareHooks from "../live-share-hooks";
import { useSharingStatus } from "../teams-js-hooks/useSharingStatus";
import { TabbedList } from "../components/TabbedList";
import { LiveShareHost, app, meeting } from "@microsoft/teams-js";
import { TestLiveShareHost } from "@microsoft/live-share";
import { IN_TEAMS } from "../constants";
import { LiveShareProvider } from "@microsoft/live-share-react";

const SidePanel: FC = () => {
    const context = useTeamsContext();
    const hostRef = useRef(
        IN_TEAMS ? LiveShareHost.create() : TestLiveShareHost.create()
    );

    return (
        <LiveShareProvider host={hostRef.current} joinOnLoad>
            <LiveSharePage context={context}>
                <SidePanelContent context={context} />
            </LiveSharePage>
        </LiveShareProvider>
    );
};

const SidePanelContent: FC<{
    context: app.Context | undefined;
}> = ({ context }) => {
    const sharingActive = useSharingStatus(context);
    const navigate = useNavigate();

    // Presence hook
    const { allUsers, localUser, localUserIsEligiblePresenter } =
        liveShareHooks.usePresence();

    const { sendNotification } = liveShareHooks.useNotifications(allUsers);

    // Take control map
    const {
        takeControl, // callback method to take control of playback
    } = liveShareHooks.useTakeControl(
        localUser,
        localUserIsEligiblePresenter,
        allUsers,
        sendNotification
    );

    // Playlist map
    const {
        playlistStarted, // boolean that is true once playlistMap listener is registered
        selectedMediaItem, // selected media item object, or undefined if unknown
        mediaItems,
        addMediaItem,
        removeMediaItem,
        selectMediaId,
    } = liveShareHooks.usePlaylist(sendNotification);

    useEffect(() => {
        if (context && playlistStarted && IN_TEAMS) {
            if (context.page?.frameContext === "meetingStage") {
                // User shared the app directly to stage, redirect automatically
                selectMediaId(mediaList[0].id);
                navigate({
                    pathname: "/",
                    search: `?inTeams=true`,
                });
            }
        }
    }, [context, playlistStarted, navigate, selectMediaId]);

    const selectMedia = useCallback(
        (mediaItem: MediaItem) => {
            // Take control
            takeControl();
            // Set the selected media ID in the playlist map
            selectMediaId(mediaItem.id);
            if (IN_TEAMS) {
                // If not already sharing to stage, share to stage
                if (!sharingActive) {
                    meeting.shareAppContentToStage((error) => {
                        if (error) {
                            console.error(error);
                        }
                    }, `${window.location.origin}/?inTeams=true`);
                }
            } else {
                // When testing locally, open in a new browser tab
                // window.open(`${window.location.origin}/`);
            }
        },
        [sharingActive, selectMediaId, takeControl]
    );

    return (
        <ListWrapper>
            <TabbedList
                mediaItems={mediaItems}
                browseItems={searchList}
                sharingActive={sharingActive}
                nowPlayingId={selectedMediaItem?.id}
                addMediaItem={addMediaItem}
                removeMediaItem={removeMediaItem}
                selectMedia={selectMedia}
            />
        </ListWrapper>
    );
};

export default SidePanel;
