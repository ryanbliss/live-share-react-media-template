import { useEffect, useRef, useState } from "react";
import { getDriveId } from "@fluidframework/odsp-doclib-utils";

export const useDriveId = (
    siteUrl: string,
    spoToken: string | undefined
): {
    driveId: string | undefined;
    loading: boolean;
    error: Error | undefined;
} => {
    const [driveId, setDriveId] = useState<string>();
    const [error, setError] = useState<Error>();
    const startedRef = useRef(false);

    useEffect(() => {
        if (startedRef.current) return;
        if (!spoToken) return;
        startedRef.current = true;
        getDriveId(siteUrl, "", undefined, {
            accessToken: spoToken,
        })
            .then((value) => {
                setDriveId(value);
            })
            .catch((err: unknown) => {
                if (err instanceof Error) {
                    setError(err);
                } else {
                    setError(new Error("Unable to get driveId"));
                }
            });
        return () => {
            startedRef.current = false;
        };
    }, [siteUrl, spoToken]);

    return {
        driveId,
        error,
        loading: !driveId && !error,
    };
};
