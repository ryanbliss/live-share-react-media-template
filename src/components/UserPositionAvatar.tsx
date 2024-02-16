import { FC } from "react";
import { FollowModeUser, TLiveFollowMode } from "../live-share-hooks";
import { AzureMediaPlayer } from "../utils/AzureMediaPlayer";
import { IPlayerState } from "./MediaPlayerContainer";
import { useLiveShareContext } from "@microsoft/live-share-react";
import { Avatar } from "@fluentui/react-components";

export interface IUserPositionAvatarProps {
    user: FollowModeUser;
    playerState: IPlayerState;
    liveFollowMode: TLiveFollowMode;
    onClick: () => void;
}

export const ColorBrand = () => (
    <Avatar color="brand" initials="BR" name="brand color avatar" />
);

export const Name = () => <Avatar name="Ashley McCarthy" />;

export const UserPositionAvatar: FC<IUserPositionAvatarProps> = ({
    user,
    playerState,
    liveFollowMode,
    onClick,
}) => {
    const { timestampProvider } = useLiveShareContext();
    if (!timestampProvider || !liveFollowMode) return null;
    const followingUser = user.data?.followingUserId
        ? liveFollowMode.getUser(user.data?.followingUserId)
        : user;
    if (!followingUser) return null;
    const paused = followingUser.data?.stateValue.paused ?? true;
    const mediaPosition =
        followingUser.data?.stateValue?.changed?.mediaPosition ?? 0;
    const timestamp = followingUser.data?.stateValue?.changed?.timestamp ?? 0;
    const percent = paused
        ? mediaPosition / playerState.duration
        : Math.min(
              (timestampProvider.getTimestamp() -
                  timestamp +
                  mediaPosition * 1000) /
                  1000 /
                  playerState.duration,
              1
          );

    return (
        <Avatar
            color="brand"
            name={user.displayName}
            size={20}
            style={{
                position: "relative",
                left: `${percent * 100}%`,
                color: "white",
                bottom: "4px",
                cursor: "pointer",
            }}
            onClick={onClick}
        />
    );
};
