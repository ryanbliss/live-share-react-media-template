/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import {
    ExtendedMediaMetadata,
    IMediaPlayerSynchronizerEvent,
    MediaPlayerSynchronizerEvents,
} from "@microsoft/live-share-media";
import { useEffect, useCallback } from "react";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";
import { MediaItem } from "../utils/media-list";
import {
    useLiveShareContext,
    useMediaSynchronizer,
} from "@microsoft/live-share-react";
import {
    ACCEPT_PLAYBACK_CHANGES_FROM,
    AppConfiguration,
    IN_TEAMS,
    UNIQUE_KEYS,
} from "../constants";
import { meeting } from "@microsoft/teams-js";
import { DisplayNotificationCallback } from "./useNotifications";
import { UpdateFollowStateCallback } from "./useTakeControl";

/**
 * Hook that synchronizes a media element using MediaSynchronizer and LiveMediaSession
 *
 * @remarks
 * Works with any HTML5 <video> or <audio> element.
 * Must use custom media controls to intercept play, pause, and seek events.
 * Any pause/play/seek events not sent through the MediaSynchronizer will be blocked
 * while MediaSynchronizer is synchronizing.
 */
export const useMediaSession = (
    keyPrefix: string,
    localUserIsPresenting: boolean,
    isShareInitiator: boolean,
    player: AzureMediaPlayer | null,
    selectedMediaItem: MediaItem | undefined,
    displayNotification: DisplayNotificationCallback,
    updateFollowState: UpdateFollowStateCallback
) => {
    const { timestampProvider } = useLiveShareContext();
    const canSendPositionUpdates = AppConfiguration.isFullyLargeMeetingOptimized
        ? localUserIsPresenting || isShareInitiator
        : true;
    const { mediaSynchronizer, suspended, beginSuspension, endSuspension } =
        useMediaSynchronizer(
            `${keyPrefix}/${UNIQUE_KEYS.media}`, // unique key for meeting + media
            player,
            selectedMediaItem?.src ?? null,
            ACCEPT_PLAYBACK_CHANGES_FROM,
            !localUserIsPresenting, // viewOnly for can play/pause/seek/setTrack
            canSendPositionUpdates // canSendPositionUpdates for large meeting optimizations
        );

    // callback method to change the selected track src
    const setTrack = useCallback(
        async (trackId: string) => {
            if (!localUserIsPresenting) return;
            updateFollowState({
                mediaId: selectedMediaItem?.id,
                paused: true,
                changed: {
                    timestamp: timestampProvider?.getTimestamp() ?? 0,
                    mediaPosition: 0,
                },
            });
            const metadata: ExtendedMediaMetadata = {
                trackIdentifier: trackId,
                liveStream: false,
                album: "",
                artist: "",
                artwork: [],
                title: selectedMediaItem ? selectedMediaItem?.title : "",
            };
            mediaSynchronizer?.setTrack(metadata);
        },
        [
            mediaSynchronizer,
            selectedMediaItem,
            localUserIsPresenting,
            updateFollowState,
        ]
    );

    // callback method to play through the synchronizer
    const play = useCallback(async () => {
        updateFollowState({
            mediaId: selectedMediaItem?.id,
            paused: false,
            changed: {
                timestamp: timestampProvider?.getTimestamp() ?? 0,
                mediaPosition: player?.currentTime ?? 0,
            },
        });
        if (localUserIsPresenting) {
            // Synchronize the play action
            mediaSynchronizer?.play();
        } else {
            // Stop following the presenter and play
            if (!suspended) {
                // Suspends media session coordinator until suspension is ended
                beginSuspension();
            }
            player?.play();
        }
    }, [
        mediaSynchronizer,
        selectedMediaItem,
        localUserIsPresenting,
        player,
        suspended,
        timestampProvider,
        beginSuspension,
        endSuspension,
        updateFollowState,
    ]);

    // callback method to play through the synchronizer
    const pause = useCallback(async () => {
        updateFollowState({
            mediaId: selectedMediaItem?.id,
            paused: true,
            changed: {
                timestamp: timestampProvider?.getTimestamp() ?? 0,
                mediaPosition: player?.currentTime ?? 0,
            },
        });
        if (localUserIsPresenting) {
            // Synchronize the pause action
            mediaSynchronizer?.pause();
        } else {
            // Stop following the presenter and pause
            if (!suspended) {
                // Suspends media session coordinator until suspension is ended
                beginSuspension();
            }
            player?.pause();
        }
    }, [
        mediaSynchronizer,
        selectedMediaItem,
        localUserIsPresenting,
        player,
        suspended,
        beginSuspension,
        endSuspension,
        updateFollowState,
    ]);

    // callback method to seek a video to a given timestamp (in seconds)
    const seekTo = useCallback(
        async (timestamp: number) => {
            updateFollowState({
                mediaId: selectedMediaItem?.id,
                paused: player?.paused ?? true,
                changed: {
                    timestamp: timestampProvider?.getTimestamp() ?? 0,
                    mediaPosition: timestamp,
                },
            });
            if (localUserIsPresenting) {
                // Synchronize the seek action
                mediaSynchronizer?.seekTo(timestamp);
            } else {
                // Stop following the presenter and seek
                if (!suspended) {
                    // Suspends media session coordinator until suspension is ended
                    beginSuspension();
                }
                if (player) {
                    player.currentTime = timestamp;
                }
            }
        },
        [
            mediaSynchronizer,
            selectedMediaItem,
            localUserIsPresenting,
            player,
            suspended,
            beginSuspension,
            endSuspension,
            updateFollowState,
        ]
    );

    // Hook to set player to view only mode when user is not the presenter and set track if needed
    useEffect(() => {
        if (!mediaSynchronizer) return;
        if (!localUserIsPresenting) return;
        if (!selectedMediaItem) return;
        const currentSrc = mediaSynchronizer.player.src;
        if (currentSrc && currentSrc === selectedMediaItem?.src) return;
        setTrack(selectedMediaItem.src);
    }, [
        localUserIsPresenting,
        mediaSynchronizer,
        selectedMediaItem?.src,
        setTrack,
    ]);

    // Register audio ducking
    useEffect(() => {
        if (!mediaSynchronizer) return;
        if (!IN_TEAMS) return;
        // Will replace existing handler
        meeting.registerSpeakingStateChangeHandler((speakingState) => {
            if (speakingState.isSpeakingDetected) {
                mediaSynchronizer?.volumeManager?.startLimiting();
            } else {
                mediaSynchronizer?.volumeManager?.stopLimiting();
            }
        });
    }, [mediaSynchronizer]);

    // Register group event listener for display notifications
    useEffect(() => {
        if (!mediaSynchronizer) return;
        function onGroupAction(evt: IMediaPlayerSynchronizerEvent): void {
            if (!mediaSynchronizer) return;
            if (evt.details.source !== "user") return;
            let displayText: string;
            switch (evt.details.action) {
                case "play": {
                    displayText = `played the ${selectedMediaItem?.type}`;
                    break;
                }
                case "pause": {
                    displayText = `paused the ${selectedMediaItem?.type}`;
                    break;
                }
                case "seekto": {
                    displayText = `seeked the ${selectedMediaItem?.type}`;
                    break;
                }
                case "settrack": {
                    displayText = `changed the ${selectedMediaItem?.type}`;
                    break;
                }
                default: {
                    return;
                }
            }
            displayNotification(
                mediaSynchronizer.mediaSession,
                displayText,
                evt.details.clientId,
                localUserIsPresenting
            );
        }
        mediaSynchronizer.addEventListener(
            MediaPlayerSynchronizerEvents.groupaction,
            onGroupAction
        );
        return () => {
            mediaSynchronizer.removeEventListener(
                MediaPlayerSynchronizerEvents.groupaction,
                onGroupAction
            );
        };
    }, [
        mediaSynchronizer,
        selectedMediaItem,
        localUserIsPresenting,
        displayNotification,
    ]);

    // Return relevant objects and callbacks UI layer
    return {
        mediaSessionStarted: !!mediaSynchronizer,
        suspended,
        play,
        pause,
        seekTo,
        setTrack,
        endSuspension,
    };
};
