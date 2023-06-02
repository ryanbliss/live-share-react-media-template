/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { RefObject, useEffect } from "react";
import { UNIQUE_KEYS } from "../constants";
import { useLiveCanvas } from "@microsoft/live-share-react";

/**
 * Sets up LiveCanvas instance
 */
export const useInkingManager = (
    hostingElement: RefObject<HTMLDivElement | null>
) => {
    const { inkingManager, liveCanvas } = useLiveCanvas(
        UNIQUE_KEYS.inking,
        hostingElement
    );

    useEffect(() => {
        inkingManager?.activate();
    }, [inkingManager]);

    return {
        inkingManager,
        liveCanvas,
    };
};
