/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { useEffect, useState, useRef, FC } from "react";
import * as liveShareHooks from "../live-share-hooks";
import {
    LiveNotifications,
    LiveSharePage,
    MediaPlayerContainer,
} from "../components";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";
import { useTeamsContext } from "../teams-js-hooks/useTeamsContext";
import { LiveShareProvider } from "@microsoft/live-share-react";
import { IN_TEAMS } from "../constants";
import { LiveShareHost } from "@microsoft/teams-js";
import { TestLiveShareHost } from "@microsoft/live-share";

const MeetingStage: FC = () => {
    // Teams context
    const context = useTeamsContext();

    const hostRef = useRef(
        IN_TEAMS ? LiveShareHost.create() : TestLiveShareHost.create()
    );

    // Render the media player
    return (
        <LiveShareProvider host={hostRef.current} joinOnLoad>
            <div style={{ backgroundColor: "black" }}>
                {/* Live Share wrapper to show loading indicator before setup */}
                <LiveSharePage context={context}>
                    <MeetingStageContent />
                </LiveSharePage>
            </div>
        </LiveShareProvider>
    );
};

const MeetingStageContent: FC = () => {
    // Element ref for inking canvas
    const canvasRef = useRef<HTMLDivElement | null>(null);
    // Media player
    const [player, setPlayer] = useState<AzureMediaPlayer | null>(null);
    // Flag tracking whether player setup has started
    const playerSetupStarted = useRef(false);

    // Presence hook
    const { allUsers, localUser, localUserIsEligiblePresenter } =
        liveShareHooks.usePresence();

    const { notificationToDisplay, sendNotification } =
        liveShareHooks.useNotifications(allUsers);

    // Take control map
    const {
        localUserIsPresenting, // boolean that is true if local user is currently presenting
        takeControl, // callback method to take control of playback
    } = liveShareHooks.useTakeControl(
        localUser,
        localUserIsEligiblePresenter,
        allUsers,
        sendNotification
    );

    // Playlist map
    const {
        selectedMediaItem, // selected media item object, or undefined if unknown
        nextTrack, // callback method to skip to the next track
    } = liveShareHooks.usePlaylist(sendNotification);

    // Media session hook
    const {
        suspended, // boolean that is true if synchronizer is suspended
        play, // callback method to synchronize a play action
        pause, // callback method to synchronize a pause action
        seekTo, // callback method to synchronize a seekTo action
        endSuspension, // callback method to end the synchronizer suspension
    } = liveShareHooks.useMediaSession(
        localUserIsPresenting,
        player,
        selectedMediaItem,
        sendNotification
    );

    // useLiveCanvas hook will insert the canvas as a child of hosting element
    // and starts the Live Inking session.It returns set of callbacks for clearing
    // the canvas, changing Ink tool type, and brush colors.
    const {
        inkingManager, // Manager class
        liveCanvas, // LiveCanvas instance
    } = liveShareHooks.useInkingManager(canvasRef);

    // Set up the media player
    useEffect(() => {
        if (player || !selectedMediaItem || playerSetupStarted.current) return;
        playerSetupStarted.current = true;
        // Setup Azure Media Player
        const amp = new AzureMediaPlayer("video", selectedMediaItem.src);
        // Set player when AzureMediaPlayer is ready to go
        const onReady = () => {
            setPlayer(amp);
            amp.removeEventListener("ready", onReady);
        };
        amp.addEventListener("ready", onReady);
    }, [selectedMediaItem, player, setPlayer]);

    return (
        <>
            {/* Display Notifications */}
            <LiveNotifications notificationToDisplay={notificationToDisplay} />
            {/* Media Player */}
            <MediaPlayerContainer
                player={player}
                localUserIsPresenting={localUserIsPresenting}
                localUserIsEligiblePresenter={localUserIsEligiblePresenter}
                suspended={suspended}
                canvasRef={canvasRef}
                inkingManager={inkingManager}
                liveCanvas={liveCanvas}
                play={play}
                pause={pause}
                seekTo={seekTo}
                takeControl={takeControl}
                endSuspension={endSuspension}
                nextTrack={nextTrack}
            >
                {/* // Render video */}
                <video
                    id="video"
                    className="azuremediaplayer amp-default-skin amp-big-play-centered"
                />
            </MediaPlayerContainer>
        </>
    );
};

export default MeetingStage;
