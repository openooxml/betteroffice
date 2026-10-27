// Re-export the framework-agnostic helpers so existing React imports keep
// working; the canonical implementations live in core.
export {
  getCommentText,
  formatDate,
  getInitials,
  getAvatarColor,
  truncateText,
  type TrackedChangeEntry,
} from '@betteroffice/docx/utils/comments';
