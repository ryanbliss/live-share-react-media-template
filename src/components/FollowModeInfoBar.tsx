import { FC, ReactNode } from "react";
import { FlexRow } from "./flex";
import { FollowModeType, IFollowModeState } from "@microsoft/live-share";
import { tokens } from "@fluentui/react-theme";
import { FollowModeState } from "../live-share-hooks";

interface IFollowModeInfoBarProps {
    children?: ReactNode;
    followState: FollowModeState;
}

export const FollowModeInfoBar: FC<IFollowModeInfoBarProps> = ({
    children,
    followState,
}) => {
    return (
        <FlexRow
            hAlign="center"
            vAlign="center"
            style={{
                position: "fixed",
                top: 12,
                left: "50%",
                transform: "translate(-50% , 0%)",
                "-webkit-transform": "translate(-50%, 0%)",
                paddingBottom: "4px",
                paddingTop: "4px",
                paddingLeft: "16px",
                paddingRight: "4px",
                borderRadius: "4px",
                minHeight: "24px",
                backgroundColor:
                    followState?.type === FollowModeType.activePresenter ||
                    followState?.type === FollowModeType.activeFollowers
                        ? tokens.colorPaletteRedBackground3
                        : tokens.colorPaletteBlueBorderActive,
            }}
        >
            {children}
        </FlexRow>
    );
};
