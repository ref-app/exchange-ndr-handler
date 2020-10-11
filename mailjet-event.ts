import * as ews from "ews-javascript-api";

// https://dev.mailjet.com/email/guides/webhooks/#event-types
type MailjetMessageStatus =
  | "unknown"
  | "queued"
  | "sent"
  | "opened"
  | "clicked"
  | "bounce"
  | "spam"
  | "unsub"
  | "blocked"
  | "hardbounced"
  | "softbounced"
  | "deferred";

type MailjetErrorRelatedTo =
  | "recipient"
  | "mailbox inactive"
  | "quota exceeded"
  | "blacklisted"
  | "spam reporter"
  | "domain"
  | "no mail host"
  | "relay/access denied"
  | "greylisted"
  | "typofix"
  | "content"
  | "error in template language"
  | "spam"
  | "content blocked"
  | "policy issue"
  | "system"
  | "protocol issue"
  | "connection issue"
  | "mailjet"
  | "duplicate in campaign";

interface MailjetEvent {
  event: MailjetMessageStatus;
  time: number;
  /**
   * Bigint in Mailjet API, always string here
   */
  MessageID: string | number;
  Message_GUID: string;
  email?: string;
  mj_campaign_id?: number;
  mj_contact_id?: number;
  customcampaign?: string;
  mj_message_id?: string;
  smtp_reply?: string;
  CustomID?: string;
  Payload?: string;
  blocked?: boolean;
  hard_bounce?: boolean;
  error_related_to: MailjetErrorRelatedTo;
  error: string;
}

export function createMailjetEvent(
  ndrItem: ews.Item,
  originalItem: ews.Item,
  errorCode: string
) {}
