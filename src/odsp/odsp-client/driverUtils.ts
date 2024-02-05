import {
    ISummaryTree,
    SummaryType,
} from "@fluidframework/protocol-definitions";

/**
 * Defines the current layout of an .app + .protocol summary tree
 * this is used internally for create new, and single commit summary
 * @internal
 */
export interface CombinedAppAndProtocolSummary extends ISummaryTree {
    tree: {
        [".app"]: ISummaryTree;
        [".protocol"]: ISummaryTree;
    };
}

/**
 * Validates the current layout of an .app + .protocol summary tree
 * this is used internally for create new, and single commit summary
 * @internal
 */
export function isCombinedAppAndProtocolSummary(
    summary: ISummaryTree | undefined,
    ...optionalRootTrees: string[]
): summary is CombinedAppAndProtocolSummary {
    if (
        summary?.tree === undefined ||
        summary.tree?.[".app"]?.type !== SummaryType.Tree ||
        summary.tree?.[".protocol"]?.type !== SummaryType.Tree
    ) {
        return false;
    }
    const treeKeys = Object.keys(summary.tree).filter(
        (t) => !optionalRootTrees.includes(t)
    );
    if (treeKeys.length !== 2) {
        return false;
    }
    return true;
}
