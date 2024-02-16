/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import {
    useEffect,
    useState,
    useCallback,
    FC,
    ReactNode,
    MutableRefObject,
} from "react";
import useResizeObserver from "use-resize-observer";
import PlayerProgressBar from "./PlayerProgressBar";
import { debounce } from "lodash";
import { mergeClasses, tokens } from "@fluentui/react-components";
import {
    getFlexColumnStyles,
    getFlexItemStyles,
    getFlexRowStyles,
} from "../styles/layouts";
import {
    getPlayerControlStyles,
    getResizeReferenceStyles,
    getVideoStyle,
} from "../styles/styles";
import { InkCanvas } from "./InkCanvas";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";
import { InkingManager, LiveCanvas } from "@microsoft/live-share-canvas";
import { useVisibleVideoSize } from "../utils/useVisibleVideoSize";
import { PlayerControls } from "./PlayerControls";
import {
    FollowModeState,
    FollowModeUser,
    FollowUserStateCallback,
    IFollowModeData,
    TLiveFollowMode,
    UpdateFollowStateCallback,
} from "../live-share-hooks";
import { FlexRow } from "./flex";
import { UserPositionAvatar } from "./UserPositionAvatar";
import { FollowModeType, PresenceState } from "@microsoft/live-share";
import { FollowModeInfoBar } from "./FollowModeInfoBar";
import { FollowModeInfoText } from "./FollowModeInfoText";
import { FollowModeSmallButton } from "./FollowModeSmallButton";

const events = [
    "loadstart",
    "timeupdate",
    "play",
    "playing",
    "pause",
    "ended",
    "seeked",
    "seeking",
    "volumechange",
    "emptied",
];

export interface IPlayerState {
    isPlaying: boolean;
    playbackStarted: boolean;
    duration: number;
    currentTime: number;
    muted: boolean;
    volume: number;
    currentPlaybackBitrate?: number;
    currentHeuristicProfile?: string;
    resolution?: string;
}

interface IMediaPlayerContainerProps {
    localUser: FollowModeUser | undefined;
    otherUsers: FollowModeUser[];
    player: AzureMediaPlayer | null;
    liveCanvas?: LiveCanvas;
    localUserIsPresenting: boolean;
    localUserIsEligiblePresenter: boolean;
    liveFollowMode: TLiveFollowMode;
    suspended: boolean;
    play: () => void;
    pause: () => void;
    seekTo: (time: number) => void;
    takeControl: () => void;
    stopPresenting: () => Promise<void>;
    stopFollowing: () => Promise<void>;
    followUser: FollowUserStateCallback;
    endSuspension: () => void;
    nextTrack?: () => void; // todo?
    updateFollowState: UpdateFollowStateCallback;
    canvasRef: MutableRefObject<HTMLDivElement | null>;
    inkingManager?: InkingManager;
    followState: FollowModeState;
    children: ReactNode;
}

export const MediaPlayerContainer: FC<IMediaPlayerContainerProps> = ({
    localUser,
    otherUsers,
    player,
    liveCanvas,
    localUserIsPresenting,
    localUserIsEligiblePresenter,
    suspended,
    liveFollowMode,
    play,
    pause,
    seekTo,
    takeControl,
    stopPresenting,
    stopFollowing,
    followUser,
    endSuspension,
    nextTrack,
    updateFollowState,
    canvasRef,
    inkingManager,
    followState,
    children,
}) => {
    const [showControls, setShowControls] = useState(true);
    const [inkActive, setInkActive] = useState(false);
    const [playerState, setPlayerState] = useState<IPlayerState>({
        isPlaying: false,
        playbackStarted: false,
        duration: 0,
        currentTime: 0,
        muted: false,
        volume: 1,
        currentPlaybackBitrate: undefined,
        currentHeuristicProfile: undefined,
        resolution: undefined,
    });
    const { ref: resizeRef, width = 1, height = 1 } = useResizeObserver();
    const videoSize = useVisibleVideoSize(width, height);

    const hideControls = useCallback(() => {
        setShowControls(false);
    }, [setShowControls]);
    // eslint-disable-next-line
    const debouncedHideControls = useCallback(debounce(hideControls, 2500), [
        hideControls,
    ]);

    const togglePlayPause = useCallback(() => {
        if (!player) {
            return;
        }
        if (player.paused) {
            play();
        } else {
            pause();
        }
    }, [player, play, pause]);

    const toggleMute = useCallback(() => {
        if (!player) {
            return;
        }
        player.muted = !player.muted;
    }, [player]);

    useEffect(() => {
        if (!localUserIsPresenting) {
            // Disable ink
            setInkActive(false);
        }
    }, [localUserIsPresenting, setInkActive]);

    useEffect(() => {
        const onPlayerStateUpdate = () => {
            if (!player) {
                return;
            }
            setPlayerState({
                isPlaying: !player.paused,
                playbackStarted: player.currentTime > 0,
                duration: player.duration || 0,
                currentTime: player.currentTime || 0,
                muted: player.muted,
                volume: player.volume,
                currentPlaybackBitrate: player.currentPlaybackBitrate,
                currentHeuristicProfile: player.currentHeuristicProfile,
                resolution: player.resolution,
            });
        };

        if (player) {
            // Add event listeners to player
            console.log("CustomControls: listening to player state changes");
            events.forEach((evt) => {
                player.addEventListener(evt, onPlayerStateUpdate);
            });
        }

        return () => {
            events.forEach((evt) => {
                player?.removeEventListener(evt, onPlayerStateUpdate);
            });
        };
    }, [player]);

    useEffect(() => {
        if (player && togglePlayPause) {
            document.body.onkeyup = function (e) {
                e.preventDefault();
                if (e.key === " " || e.code === "Space") {
                    togglePlayPause();
                }
            };
        }
    }, [player, togglePlayPause]);

    const flexRowStyles = getFlexRowStyles();
    const flexColumnStyles = getFlexColumnStyles();
    const flexItemStyles = getFlexItemStyles();
    const playerControlStyles = getPlayerControlStyles();
    const videoStyle = getVideoStyle();
    const resizeReferenceStyles = getResizeReferenceStyles();

    return (
        <div
            style={{
                color: tokens.colorNeutralForegroundStaticInverted,
            }}
            className={mergeClasses(
                flexColumnStyles.root,
                playerControlStyles.root
            )}
            onMouseMove={() => {
                setShowControls(true);
                debouncedHideControls();
            }}
        >
            <div className={resizeReferenceStyles.root} ref={resizeRef} />
            <div
                className={videoStyle.root}
                onClick={togglePlayPause}
                style={{
                    left: `${videoSize?.xOffset || 0}px`,
                    top: `${videoSize?.yOffset || 0}px`,
                    width: `${videoSize?.width || 0}px`,
                    height: `${videoSize?.height || 0}px`,
                }}
            >
                {children}
            </div>
            <InkCanvas
                canvasRef={canvasRef}
                isEnabled={inkActive}
                inkingManager={inkingManager}
                videoSize={videoSize}
            />
            <div
                className={flexColumnStyles.root}
                style={{
                    position: "absolute",
                    left: "0",
                    bottom: "0",
                    right: "0",
                    zIndex: 2,
                    visibility:
                        showControls || !playerState.isPlaying
                            ? "visible"
                            : "hidden",
                    background:
                        "linear-gradient(rgba(0,0,0,0), rgba(0,0,0,0.4))",
                }}
            >
                {/* Follow mode information / actions */}
                {!!followState && followState.type !== FollowModeType.local && (
                    <FollowModeInfoBar followState={followState}>
                        <FollowModeInfoText
                            localUser={localUser}
                            otherUsers={otherUsers}
                            followState={followState}
                            liveFollowMode={liveFollowMode}
                        />
                        {followState.type ===
                            FollowModeType.activePresenter && (
                            <FollowModeSmallButton onClick={stopPresenting}>
                                {"STOP"}
                            </FollowModeSmallButton>
                        )}
                        {followState.type === FollowModeType.followUser && (
                            <FollowModeSmallButton onClick={stopFollowing}>
                                {"STOP"}
                            </FollowModeSmallButton>
                        )}
                        {followState.type ===
                            FollowModeType.suspendFollowPresenter && (
                            <FollowModeSmallButton onClick={endSuspension}>
                                {"FOLLOW"}
                            </FollowModeSmallButton>
                        )}
                        {followState.type ===
                            FollowModeType.suspendFollowUser && (
                            <FollowModeSmallButton onClick={endSuspension}>
                                {"RESUME"}
                            </FollowModeSmallButton>
                        )}
                    </FollowModeInfoBar>
                )}
                {/* User position avatars */}
                <div
                    style={{
                        position: "absolute",
                        color: "white",
                        left: 0,
                        right: 0,
                    }}
                >
                    {otherUsers
                        .filter(
                            (user) =>
                                user.state !== PresenceState.offline &&
                                user.data?.followingUserId !==
                                    localUser?.userId &&
                                user.userId !== followState?.followingUserId &&
                                (localUser?.data?.followingUserId ||
                                user.data?.followingUserId
                                    ? localUser?.data?.followingUserId !==
                                      user.data?.followingUserId
                                    : true)
                        )
                        .map((user) => (
                            <UserPositionAvatar
                                key={user.userId}
                                user={user}
                                playerState={playerState}
                                liveFollowMode={liveFollowMode}
                                onClick={() => {
                                    console.log("following user");
                                    followUser(
                                        user.data?.followingUserId ??
                                            user.userId
                                    )
                                        .then(() => {
                                            if (!user.data) return;
                                            updateFollowState(
                                                user.data.stateValue
                                            );
                                        })
                                        .catch((err) => console.error(err));
                                }}
                            />
                        ))}
                </div>
                {/* Seek Progress Bar */}
                <PlayerProgressBar
                    currentTime={playerState.currentTime}
                    duration={playerState.duration}
                    isPlaybackDisabled={!playerState.playbackStarted}
                    onSeek={seekTo}
                />
                <PlayerControls
                    endSuspension={endSuspension}
                    inkActive={inkActive}
                    inkingManager={inkingManager}
                    liveCanvas={liveCanvas}
                    localUserIsEligiblePresenter={localUserIsEligiblePresenter}
                    localUserIsPresenting={
                        followState?.type === FollowModeType.activePresenter
                    }
                    nextTrack={nextTrack}
                    playerState={playerState}
                    setInkActive={setInkActive}
                    suspended={suspended}
                    takeControl={takeControl}
                    toggleMute={toggleMute}
                    togglePlayPause={togglePlayPause}
                />
            </div>
        </div>
    );
};
