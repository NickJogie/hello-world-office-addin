/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

/**
 * Insert a blue paragraph in word when the add-in command is executed.
 * @param event
 */
/*
export async function insertBlueParagraphInWord(event: Office.AddinCommands.Event) {
  try {
    await Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      paragraph.font.color = "blue";
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
*/
import 'whatwg-fetch';

const summarizeWithAzureOpenAI = async (text: string): Promise<string> => {
  const apiKey = ""; 
  const endpoint = "";
  const deploymentName = "";
  const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2024-10-21`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": apiKey,
    },
    body: JSON.stringify({
      messages: [
        { role: "system", content: "You summarize text in a concise sentence." },
        { role: "user", content: `Summarize this:\n\n${text}` },
      ],
      max_tokens: 100,
      temperature: 0.3,
    }),
  });

  const data = await response.json();
  return data.choices?.[0]?.message?.content ?? "No summary generated.";
};

// Word add-in command handler
export async function insertBlueParagraphInWord(event: Office.AddinCommands.Event) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");

    await context.sync();

    const inputText = selection.text.trim();
    if (!inputText) {
      console.error("No text selected.");
      event.completed();
      return;
    }

    try {
      const summary = await summarizeWithAzureOpenAI(inputText);
      const shortSummary = summary.length > 150 ? summary.substring(0, 147) + "..." : summary;

      // Insert the summary after the selection
      selection.insertText(`\n\nSummary: ${shortSummary}`, Word.InsertLocation.end);
    } catch (error) {
      console.error("Error during summarization:", error);
      selection.insertText("\n\n[Error generating summary]", Word.InsertLocation.end);
    }

    await context.sync();
    event.completed();
  });
}
