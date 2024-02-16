import { FC } from "react";
import { Text, tokens } from "@fluentui/react-components";
import { FollowModeType } from "@microsoft/live-share";
import {
    FollowModeState,
    FollowModeUser,
    TLiveFollowMode,
} from "../live-share-hooks";

export const FollowModeInfoText: FC<{
    localUser: FollowModeUser | undefined;
    otherUsers: FollowModeUser[];
    followState: FollowModeState;
    liveFollowMode: TLiveFollowMode;
}> = ({ localUser, otherUsers, followState, liveFollowMode }) => {
    // Get the list of users following the current followed user
    const followers =
        liveFollowMode && followState?.followingUserId
            ? liveFollowMode.getUserFollowers(followState.followingUserId)
            : [];
    const localFollowers =
        liveFollowMode &&
        localUser &&
        followState?.type === FollowModeType.activeFollowers
            ? liveFollowMode.getUserFollowers(localUser.userId)
            : [];
    const followingUser = followState?.followingUserId
        ? liveFollowMode?.getUser(followState.followingUserId)
        : undefined;
    function getTextToDisplay(): string {
        if (!followState) {
            throw new Error(
                "FollowModeInfoText getTextInfoDisplay(): this function should not be called if remoteCameraState is null"
            );
        }
        switch (followState.type) {
            case FollowModeType.activePresenter: {
                if (otherUsers.length !== 1) {
                    return `Presenting to ${otherUsers.length} others`;
                }
                const nonLocalUser = otherUsers[0];
                return `Presenting to ${nonLocalUser.displayName}`;
            }
            case FollowModeType.activeFollowers: {
                if (localFollowers.length === 1) {
                    return `${localFollowers[0].displayName} is following you`;
                }
                return `${localFollowers.length} others are following you`;
            }
            case FollowModeType.followPresenter:
            case FollowModeType.suspendFollowPresenter: {
                return `${followingUser?.displayName} is presenting`;
            }
            case FollowModeType.followUser: {
                if (followers.length > 1) {
                    `You + ${followers.length - 1} others are following ${
                        followingUser?.displayName
                    }`;
                }
                return `You are following ${followingUser?.displayName}`;
            }
            case FollowModeType.suspendFollowUser: {
                return `Paused following ${followingUser?.displayName}`;
            }
            default:
                return "Invalid FollowModeType";
        }
    }
    if (!followState) return null;

    return (
        <Text
            align="center"
            style={{
                color: tokens.colorNeutralForegroundOnBrand,
                marginRight: "12px",
            }}
        >
            {getTextToDisplay()}
        </Text>
    );
};
