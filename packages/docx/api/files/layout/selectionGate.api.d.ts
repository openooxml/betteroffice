// @betteroffice/docx declaration layout/selectionGate.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Coordinates selection paint with the latest completed layout sequence. */
export declare class LayoutSelectionGate {
    #private;
    setStateSeq(sequence: number): void;
    incrementStateSeq(): number;
    getStateSeq(): number;
    getRenderSeq(): number;
    onLayoutStart(): void;
    onLayoutComplete(sequence: number): void;
    isSafeToRender(): boolean;
    requestRender(): void;
    onRender(callback: () => void): () => void;
    reset(): void;
    getDebugInfo(): {
        stateSeq: number;
        renderSeq: number;
        layoutUpdating: boolean;
        hasPendingRender: boolean;
        isSafe: boolean;
    };
}
