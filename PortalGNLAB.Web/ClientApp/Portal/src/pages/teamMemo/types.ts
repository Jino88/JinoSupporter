export interface Folder {
  id: string;
  parentId: string | null;
  name: string;
  path: string;
  ordinal: number;
  isPlan: boolean;
  isHidden: boolean;
}
export interface Attachment {
  id: string;
  fileName: string;
  contentType: string;
  size: number;
  author: string;
  createdAt: string;
  isImage: boolean;
}
export interface Comment {
  id: number;
  author: string;
  body: string;
  isResult: boolean;
  createdAt: string;
  updatedAt: string;
  attachments: Attachment[];
}
export interface Summary {
  id: string;
  folderId: string | null;
  title: string;
  kind: string;
  status: string;
  owner: string;
  tags: string;
  startDate: string | null;
  dueDate: string | null;
  isPinned: boolean;
  updatedBy: string;
  updatedAt: string;
  createdAt: string;
  entryCount: number;
  commentCount: number;
  participants: string[];
  lastCommentBy: string | null;
  lastCommentBody: string | null;
  lastCommentIsResult: boolean;
  lastActivityAt: string | null;
  repeatWeekdays: string;
  repeatUntil: string | null;
  scheduleStartAt: string | null;
  scheduleEndAt: string | null;
  scheduleAllDay: boolean;
}
export interface AiSummary {
  aiSummary: string | null;
  aiSummaryAt: string | null;
  aiSummaryBy: string;
  aiSummaryModel: string;
  isAiSummaryStale: boolean;
}
export interface Entry extends AiSummary {
  id: string;
  memoId: string;
  entryDate: string;
  title: string;
  body: string;
  targetDate: string | null;
  doneDate: string | null;
  depth: number;
  author: string;
  createdAt: string;
  updatedAt: string;
  comments: Comment[];
  attachments: Attachment[];
}
export interface Memo extends AiSummary {
  id: string;
  folderId: string | null;
  title: string;
  kind: string;
  status: string;
  owner: string;
  tags: string;
  startDate: string | null;
  dueDate: string | null;
  description: string;
  isPinned: boolean;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
  revision: number;
  repeatWeekdays: string;
  repeatUntil: string | null;
  entries: Entry[];
  comments: Comment[];
  attachments: Attachment[];
}
export interface MemoView {
  folders: Folder[];
  memos: Summary[];
  currentUser: string;
  isAdmin: boolean;
  users: {
    username: string;
    displayName: string;
  }[];
  aiConfigured: boolean;
  translationConfigured: boolean;
}
export const statuses = ["Todo", "InProgress", "Hold", "Done", "Canceled"];
export const statusLabel = (s: string) =>
  s === "InProgress" ? "In progress" : s === "Hold" ? "On hold" : s;
export type Translate = (text: string | null | undefined) => string;
