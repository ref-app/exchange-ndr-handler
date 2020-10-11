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

function getMailjetErrorFieldsFromErrorCode(
  errorCode: string
): Pick<MailjetEvent, "event" | "hard_bounce" | "error_related_to" | "error"> {
  const error = errorCode;
  const hard_bounce = errorCode.startsWith("5.");
  // Make more granular later
  if (hard_bounce) {
    return {
      event: "hardbounced",
      error_related_to: "recipient",
      error,
      hard_bounce,
    };
  } else {
    return {
      event: "softbounced",
      error_related_to: "recipient",
      error,
      hard_bounce,
    };
  }
}
/**
 *  Mailjet event-formatted callback content except that the message ID is always a string
 * and we send the error code in the "error" field
 */
export function createMailjetEvent(
  ndrItem: ews.Item,
  originalItem: ews.Item,
  errorCode: string
) {
  const time = ndrItem.DateTimeReceived.TotalMilliSeconds;
  const errorFields = getMailjetErrorFieldsFromErrorCode(errorCode);
  const result: MailjetEvent = {
    ...errorFields,
    time,
    MessageID: originalItem.Id.UniqueId,
    Message_GUID: "",
  };
  return result;
}
