/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { FC, useCallback, useEffect, useRef, useState } from "react";
import { useTeamsContext } from "../teams-js-hooks/useTeamsContext";
import { useNavigate } from "react-router-dom";
import { MediaItem, mediaList, searchList } from "../utils/media-list";
import {
    ListWrapper,
    LiveNotifications,
    LiveSharePage,
    MediaCard,
    MediaPlayerContainer,
} from "../components";
import * as liveShareHooks from "../live-share-hooks";
import {
    ISharingStatus,
    useSharingStatus,
} from "../teams-js-hooks/useSharingStatus";
import { TabbedList } from "../components/TabbedList";
import { LiveShareHost, app, meeting } from "@microsoft/teams-js";
import {
    ILiveShareClientOptions,
    LiveDataObject,
    TestLiveShareHost,
} from "@microsoft/live-share";
import { AppConfiguration, IN_TEAMS } from "../constants";
import { LiveShareProvider } from "@microsoft/live-share-react";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";

const ContentTab: FC = () => {
    const context = useTeamsContext();
    const hostRef = useRef(
        IN_TEAMS ? LiveShareHost.create() : TestLiveShareHost.create()
    );

    return (
        <LiveShareProvider host={hostRef.current} joinOnLoad>
            <LiveSharePage context={context}>
                <Content context={context} />
            </LiveSharePage>
        </LiveShareProvider>
    );
};

const Content: FC<{
    context: app.Context | undefined;
}> = ({ context }) => {
    const { notificationToDisplay, displayNotification } =
        liveShareHooks.useNotifications();

    const threadId =
        context?.meeting?.id ??
        context?.chat?.id ??
        context?.channel?.id ??
        "unknown";

    // Take control map
    const {
        localUser, // local user
        otherUsers, // users list
        localUserIsPresenting, // boolean that is true if local user is currently presenting
        localUserIsEligiblePresenter, // boolean that is true if the local user has the required roles to present
        selectedMediaItem, // media item user should look at
        state, // follow mode state object
        liveFollowMode,
        updateFollowState, // callback to change follow state
        takeControl, // callback method to take control of playback
        followUser, // callback to follow user
        stopFollowing, // callback to stop following
        stopPresenting, // callback to stop presenting
    } = liveShareHooks.useTakeControl(
        threadId,
        false,
        displayNotification,
        undefined
    );

    return (
        <>
            {/* Display Notifications */}
            <LiveNotifications notificationToDisplay={notificationToDisplay} />
            {/* Media Player */}
            <ListWrapper>
                {!selectedMediaItem &&
                    searchList.map((mediaItem) => (
                        <MediaCard
                            key={`browse-item-${mediaItem.id}`}
                            mediaItem={mediaItem}
                            nowPlayingId={undefined}
                            sharingActive={false}
                            buttonText="Open"
                            selectMedia={(item) => {
                                updateFollowState({
                                    mediaId: item.id,
                                    paused: false,
                                });
                            }}
                        />
                    ))}
                {!!selectedMediaItem && (
                    <CollaborativeVideoPlayer
                        localUser={localUser}
                        otherUsers={otherUsers}
                        threadId={threadId}
                        selectedMedia={selectedMediaItem}
                        followState={state}
                        localUserIsPresenting={localUserIsPresenting}
                        localUserIsEligiblePresenter={
                            localUserIsEligiblePresenter
                        }
                        liveFollowMode={liveFollowMode}
                        displayNotification={displayNotification}
                        updateFollowState={updateFollowState}
                        takeControl={takeControl}
                        followUser={followUser}
                        stopPresenting={stopPresenting}
                        stopFollowing={stopFollowing}
                    />
                )}
            </ListWrapper>
        </>
    );
};

const CollaborativeVideoPlayer: FC<{
    localUser: liveShareHooks.FollowModeUser | undefined;
    otherUsers: liveShareHooks.FollowModeUser[];
    threadId: string;
    followState: liveShareHooks.FollowModeState;
    selectedMedia: MediaItem;
    localUserIsPresenting: boolean;
    localUserIsEligiblePresenter: boolean;
    liveFollowMode: liveShareHooks.TLiveFollowMode;
    displayNotification: liveShareHooks.DisplayNotificationCallback;
    updateFollowState: (
        stateValue: liveShareHooks.IFollowModeData
    ) => Promise<void>;
    takeControl: () => void;
    stopPresenting: () => Promise<void>;
    stopFollowing: () => Promise<void>;
    followUser: liveShareHooks.FollowUserStateCallback;
}> = (props) => {
    // Element ref for inking canvas
    const canvasRef = useRef<HTMLDivElement | null>(null);
    // Media player
    const [player, setPlayer] = useState<AzureMediaPlayer | null>(null);
    // Flag tracking whether player setup has started
    const playerSetupStarted = useRef(false);
    const followingUserId = props.followState?.followingUserId ?? "default";
    // Media session hook
    const {
        suspended, // boolean that is true if synchronizer is suspended
        play, // callback method to synchronize a play action
        pause, // callback method to synchronize a pause action
        seekTo, // callback method to synchronize a seekTo action
        endSuspension, // callback method to end the synchronizer suspension
    } = liveShareHooks.useMediaSession(
        `${props.threadId}/${followingUserId}`,
        props.localUserIsPresenting,
        false,
        player,
        props.selectedMedia,
        props.displayNotification,
        props.updateFollowState
    );

    // Set up the media player
    useEffect(() => {
        if (player || playerSetupStarted.current) return;
        playerSetupStarted.current = true;
        // Setup Azure Media Player
        const amp = new AzureMediaPlayer("video", props.selectedMedia.src);
        // Set player when AzureMediaPlayer is ready to go
        const onReady = () => {
            setPlayer(amp);
            amp.removeEventListener("ready", onReady);
        };
        amp.addEventListener("ready", onReady);
    }, [player, setPlayer, props.selectedMedia.src]);

    return (
        <MediaPlayerContainer
            localUser={props.localUser}
            otherUsers={props.otherUsers}
            player={player}
            localUserIsPresenting={props.localUserIsPresenting}
            localUserIsEligiblePresenter={props.localUserIsEligiblePresenter}
            suspended={suspended}
            canvasRef={canvasRef}
            followState={props.followState}
            liveFollowMode={props.liveFollowMode}
            play={play}
            pause={pause}
            seekTo={seekTo}
            takeControl={props.takeControl}
            endSuspension={endSuspension}
            updateFollowState={props.updateFollowState}
            followUser={props.followUser}
            stopFollowing={props.stopFollowing}
            stopPresenting={props.stopPresenting}
        >
            {/* // Render video */}
            <video
                id="video"
                className="azuremediaplayer amp-default-skin amp-big-play-centered"
            />
        </MediaPlayerContainer>
    );
};

export default ContentTab;
