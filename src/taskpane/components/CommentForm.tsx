import * as React from "react";
import { useState, useEffect } from "react";
import { MentionsInput, Mention } from "react-mentions";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";

// SharePoint configuration
const siteId = "8314c8ba-c25a-4a02-bf25-d6238949ac8f";
const listId = "5f59364d-9808-4d26-8e04-2527b4fc403e";

const msalInstance = new PublicClientApplication(msalConfig);

const CommentForm: React.FC = () => {
  const [comment, setComment] = useState<string>("");
  const [commentHistory, setCommentHistory] = useState<any[]>([]);
  const [people, setPeople] = useState<any[]>([]);
  const [mentionedEmails, setMentionedEmails] = useState<string[]>([]);
  const [conversationId, setConversationId] = useState<string>("");
  const [itemId, setItemId] = useState<string>("");

  // Attempt to get conversationId from the email body (for forwarded emails)
  useEffect(() => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyContent = result.value;
        const regex = /CONVERSATION_ID:([a-zA-Z0-9\-]+)/;
        const match = bodyContent.match(regex);

        if (match && match[1]) {
          setConversationId(match[1]);
        } else if (Office.context.mailbox?.item?.conversationId) {
          setConversationId(Office.context.mailbox.item.conversationId);
        }

        if (Office.context.mailbox?.item?.itemId) {
          setItemId(Office.context.mailbox.item.itemId);
        }
      }
    });
  }, []);

  useEffect(() => {
    if (conversationId) {
      fetchCommentsFromSharePoint();
    }
  }, [conversationId]);

  const getAccessToken = async (): Promise<string> => {
    await msalInstance.initialize();
    let accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      await msalInstance.loginPopup({
        scopes: ["Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      });
      accounts = msalInstance.getAllAccounts();
    }

    const response = await msalInstance.acquireTokenSilent({
      scopes: ["Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      account: accounts[0],
    });

    return response.accessToken;
  };

  const fetchUsers = async () => {
    try {
      const token = await getAccessToken();
      const response = await fetch("https://graph.microsoft.com/v1.0/users?$top=50", {
        headers: { Authorization: `Bearer ${token}` },
      });

      const data = await response.json();
      const usersData = data.value.map((user: any) => ({
        id: user.mail || user.userPrincipalName,
        display: user.displayName,
        email: user.mail || user.userPrincipalName,
      }));

      setPeople(usersData);
    } catch (error) {
      console.error("Error fetching users:", error);
    }
  };

  useEffect(() => {
    fetchUsers();
  }, []);

  const fetchCommentsFromSharePoint = async () => {
    if (!conversationId) return;

    const token = await getAccessToken();
    const emailIdValue = encodeURIComponent(conversationId);

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$filter=fields/EmailID eq '${emailIdValue}'&$orderby=createdDateTime asc`;

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      console.error("Fetch comments failed:", await response.text());
      return;
    }

    const data = await response.json();
    const comments = data.value
      .map((item: any) => item.fields)
      .sort((a, b) => new Date(a.CreatedDate).getTime() - new Date(b.CreatedDate).getTime());

    setCommentHistory(comments);
  };

  const stripMentionsFromComment = (input: string): string => {
    return input.replace(/@\[[^\]]+\]\([^)]+\)/g, "").trim();
  };

  const extractMentionData = (input: string) => {
    const mentionRegex = /@\[([^\]]+)\]\(([^)]+)\)/g;
    const displayNames: string[] = [];
    const emails: string[] = [];

    let match;
    while ((match = mentionRegex.exec(input)) !== null) {
      displayNames.push(match[1]);
      emails.push(match[2]);
    }

    return { displayNames, emails };
  };

  const saveCommentToSharePoint = async () => {
    const token = await getAccessToken();
    const plainComment = stripMentionsFromComment(comment);
    const { displayNames, emails } = extractMentionData(comment);

    const mentionedUsersText = displayNames.join(", ");

    const fieldsData: any = {
      Title: "Email Comment",
      EmailID: conversationId,
      Comment: plainComment,
      MentionedUsers: mentionedUsersText,
      CreatedBy: Office.context.mailbox.userProfile.emailAddress,
      CreatedDate: new Date().toISOString(),
    };

    const body = { fields: fieldsData };

    await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    setMentionedEmails(emails);
  };

  const sendEmailToMentionedUsers = async () => {
    if (mentionedEmails.length === 0) {
      console.warn("No mentioned users to email.");
      return;
    }

    const token = await getAccessToken();

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const originalBody = result.value;

        const toRecipients = mentionedEmails.map((email) => ({
          emailAddress: { address: email },
        }));

        const emailPayload = {
          message: {
            subject: "You’ve been mentioned in an Outlook conversation",
            body: {
              contentType: "HTML",
              content: `
                <p>Hello,</p>
                <p>You were mentioned in a conversation. Here's the original email:</p>
                <hr />
                ${originalBody}
                <p>Open your Outlook Add-in to reply or view further comments.</p>
              `,
            },
            toRecipients,
          },
          saveToSentItems: true,
        };

        const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(emailPayload),
        });

        if (!response.ok) {
          console.error("SendMail API failed:", await response.text());
        } else {
          console.log("Email successfully sent to mentioned users.");
        }
      } else {
        console.error("Failed to retrieve email body.");
      }
    });
  };

  const handleSaveAndShare = async () => {
    if (!comment.trim()) {
      alert("Please add a comment before saving.");
      return;
    }

    await saveCommentToSharePoint();
    await fetchCommentsFromSharePoint();

    if (mentionedEmails.length > 0) {
      await sendEmailToMentionedUsers();
    }

    setComment("");
    setMentionedEmails([]);
  };

  return (
    <div style={{ padding: "1rem" }}>
      <h3>Comments History</h3>

      <div
        style={{
          marginBottom: "1rem",
          background: "#f9f9f9",
          padding: "10px",
          borderRadius: "5px",
          maxHeight: "200px",
          overflowY: "auto",
        }}
      >
        {commentHistory.length > 0 ? (
          commentHistory.map((c, index) => (
            <div
              key={index}
              style={{
                borderBottom: "1px solid #ddd",
                marginBottom: "5px",
                paddingBottom: "5px",
              }}
            >
              <div>{c.Comment}</div>
              {c.MentionedUsers && (
                <div style={{ color: "#0078D4", fontWeight: "bold" }}>
                  Mentioned: {c.MentionedUsers}
                </div>
              )}
            </div>
          ))
        ) : (
          <div style={{ color: "#999" }}>No comments yet.</div>
        )}
      </div>

      <h4>Add New Comment</h4>

      <MentionsInput
        value={comment}
        onChange={(e) => setComment(e.target.value)}
        style={{ width: "100%", minHeight: 80, border: "1px solid #ccc", padding: 8 }}
        placeholder="Type @ to mention someone..."
      >
        <Mention
          trigger="@"
          data={people}
          displayTransform={(_id, display) => `@${display}`}
          appendSpaceOnAdd
          onAdd={(id: string) => {
            if (!mentionedEmails.includes(id)) {
              setMentionedEmails([...mentionedEmails, id]);
            }
          }}
        />
      </MentionsInput>

      <button
        onClick={handleSaveAndShare}
        style={{
          marginTop: "15px",
          backgroundColor: "#0078D4",
          color: "#fff",
          padding: "10px",
          width: "100%",
        }}
      >
        Save & Notify Mentions
      </button>
    </div>
  );
};

export default CommentForm;
