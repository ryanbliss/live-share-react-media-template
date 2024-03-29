/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { ExtendedMediaMetadata } from "@microsoft/live-share-media";
import { useEffect, useCallback } from "react";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";
import { MediaItem } from "../utils/media-list";
import { useMediaSynchronizer } from "@microsoft/live-share-react";
import {
    ACCEPT_PLAYBACK_CHANGES_FROM,
    IN_TEAMS,
    UNIQUE_KEYS,
} from "../constants";
import { meeting } from "@microsoft/teams-js";

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
    localUserIsPresenting: boolean,
    player: AzureMediaPlayer | null,
    selectedMediaItem: MediaItem | undefined,
    sendNotification: (text: string) => void
) => {
    const { mediaSynchronizer, suspended, beginSuspension, endSuspension } =
        useMediaSynchronizer(
            UNIQUE_KEYS.media,
            player,
            selectedMediaItem?.src ?? null,
            ACCEPT_PLAYBACK_CHANGES_FROM,
            !localUserIsPresenting
        );

    // callback method to change the selected track src
    const setTrack = useCallback(
        async (trackId: string) => {
            if (!localUserIsPresenting) return;
            const metadata: ExtendedMediaMetadata = {
                trackIdentifier: trackId,
                liveStream: false,
                album: "",
                artist: "",
                artwork: [],
                title: selectedMediaItem ? selectedMediaItem?.title : "",
            };
            mediaSynchronizer?.setTrack(metadata);
            sendNotification(`changed the ${selectedMediaItem?.type}`);
        },
        [
            mediaSynchronizer,
            selectedMediaItem,
            localUserIsPresenting,
            sendNotification,
        ]
    );

    // callback method to play through the synchronizer
    const play = useCallback(async () => {
        if (localUserIsPresenting) {
            // Synchronize the play action
            mediaSynchronizer?.play();
            sendNotification(`played the ${selectedMediaItem?.type}`);
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
        beginSuspension,
        endSuspension,
        sendNotification,
    ]);

    // callback method to play through the synchronizer
    const pause = useCallback(async () => {
        if (localUserIsPresenting) {
            // Synchronize the pause action
            mediaSynchronizer?.pause();
            sendNotification(`paused the ${selectedMediaItem?.type}`);
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
        sendNotification,
    ]);

    // callback method to seek a video to a given timestamp (in seconds)
    const seekTo = useCallback(
        async (timestamp: number) => {
            if (localUserIsPresenting) {
                // Synchronize the seek action
                mediaSynchronizer?.seekTo(timestamp);
                sendNotification(`seeked the ${selectedMediaItem?.type}`);
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
            sendNotification,
        ]
    );

    // Hook to set player to view only mode when user is not the presenter and set track if needed
    useEffect(() => {
        if (!mediaSynchronizer) return;
        const currentSrc = mediaSynchronizer.player.src;
        if (currentSrc && currentSrc === selectedMediaItem?.src) return;
        if (selectedMediaItem) {
            setTrack(selectedMediaItem.src);
        }
    }, [
        localUserIsPresenting,
        mediaSynchronizer,
        selectedMediaItem?.src,
        setTrack,
    ]);

    // Register audio ducking
    useEffect(() => {
        if (!mediaSynchronizer || !IN_TEAMS) return;
        // Will replace existing handler
        meeting.registerSpeakingStateChangeHandler((speakingState) => {
            if (speakingState.isSpeakingDetected) {
                mediaSynchronizer?.volumeManager?.startLimiting();
            } else {
                mediaSynchronizer?.volumeManager?.stopLimiting();
            }
        });
    }, [mediaSynchronizer]);

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
