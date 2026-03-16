// Mail folder from Graph API
export interface MailFolder {
  id: string;
  displayName: string;
  unreadItemCount: number;
  totalItemCount: number;
}

// Mail message summary (for list/search results)
export interface MailMessage {
  id: string;
  subject: string;
  from: string;
  to: string[];
  receivedDateTime: string;
  bodyPreview: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
}

// Full mail message (for get_mail)
export interface MailMessageFull extends MailMessage {
  body: string;
  cc: string[];
  attachments: AttachmentMeta[];
  categories: string[];
}

// Attachment metadata (no content)
export interface AttachmentMeta {
  id: string;
  name: string;
  size: number;
  contentType: string;
}

// Auth status response
export interface AuthStatus {
  authenticated: boolean;
  userEmail: string | null;
  tokenExpires: string | null;
}
