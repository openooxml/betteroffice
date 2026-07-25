/**
 * Largest collaboration frame the relay accepts, broadcasts and retains.
 * Clients must not exceed it: the relay cannot replay what it cannot retain.
 */
export const MAX_COLLABORATION_FRAME_BYTES = 16 * 1024 * 1024;
