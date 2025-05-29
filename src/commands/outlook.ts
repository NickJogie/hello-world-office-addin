/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

/**
 * Show an outlook notification when the add-in command is executed.
 * @param event
 */
/*export function setNotificationInOutlook(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
*/
//import fetch from "node-fetch";
import 'whatwg-fetch'; //npm install whatwg-fetch


const summarizeWithAzureOpenAI = async (emailText: string): Promise<string> => {
  const apiKey = "";
  const endpoint = "";
  const deploymentName = "";
  //https://myaiservices.openai.azure.com/openai/deployments/gpt-35-turbo/chat/completions?api-version=2025-01-01-preview
  const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2024-10-21`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": apiKey,
    },
    body: JSON.stringify({
      messages: [
        { role: "system", content: "You summarize emails in a concise sentence." },
        { role: "user", content: `Summarize this email:\n\n${emailText}` },
      ],
      max_tokens: 100,
      temperature: 0.3,
    }),
  });

  const data = await response.json();
  return data.choices?.[0]?.message?.content ?? "No summary generated.";
};

export async function setNotificationInOutlook(event: Office.AddinCommands.Event) {
  try {
    // Get plain text email body
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailText = result.value;
        const emailTest = emailText?.substring(0, 200) ?? "(empty)";

        // Call Azure OpenAI
        const summary = await summarizeWithAzureOpenAI(emailTest);
        //const summary = emailText?.substring(0, 100) ?? "(empty)";

        // Truncate to fit notification limit (~150 chars)
        const shortSummary = summary.length > 150 ? summary.substring(0, 147) + "..." : summary;

        // Show the summary in an informational notification
        const message: Office.NotificationMessageDetails = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: `Summary: ${shortSummary}`,
          icon: "Icon.80x80",
          persistent: true,
        };

        Office.context.mailbox.item.notificationMessages.replaceAsync(
          "SummaryNotification",
          message,
          () => event.completed() // complete the event after async is done
        );
      } else {
        // fallback notification
        Office.context.mailbox.item.notificationMessages.replaceAsync("SummaryNotification", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "Failed to retrieve email body.",
        });
        event.completed();
      }
    });
  } catch (err) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("SummaryNotification", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Error during summarization.",
    });
    event.completed();
  }
}
