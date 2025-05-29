/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runOutlook;
    //document.getElementById("prepend").onclick= prependText;
    document.getElementById("prepend").addEventListener("click", prependText);
  }
});

export async function runOutlook() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

export async function prependText() {
  /* This snippet adds text to the beginning of the message or appointment's body. 
    
    When prepending a link in HTML markup to the body, you can disable the online link preview by setting the anchor tag's id attribute to "LPNoLP". For example, '<a id="LPNoLP" href="https://www.contoso.com">Click here!</a>'.
  */
  let logArea = document.getElementById("log-area");
  const logMessage = (message) => {
    let logElement = document.createElement("p");
    logElement.textContent = message;
    logArea.appendChild(logElement);
  };
  /*
  const text = (document.getElementById("text-field") as HTMLInputElement).value;
  const centeredText = `
      <div style="text-align: center; font-size: 20px; font-weight: bold; margin-bottom: 20px;">
        ${text}
      </div>
    `;

  logMessage("Attempting to prepend text: " + text);
  */
  // Get all checked checkboxes
  const selectedValues = [];
  const checkboxes = document.querySelectorAll('#checkbox-container input[type="checkbox"]:checked');
  
  checkboxes.forEach((checkbox) => {
    // Cast checkbox to HTMLInputElement to access the 'value' property
    selectedValues.push((checkbox as HTMLInputElement).value);
  });

  if (selectedValues.length === 0) {
    logMessage("No options selected.");
    return;
  }

  logMessage("Attempting to prepend selected text: " + selectedValues.join(", "));
  // Construct the HTML content with the selected checkboxes, centered.
  const contentToPrepend = selectedValues.map(value => `
    <div style="text-align: center; font-size: 20px; font-weight: bold; margin-bottom: 20px;">
      ${value}
    </div>
  `).join("");


  // It's recommended to call getTypeAsync and pass its returned value to the options.coercionType parameter of the prependAsync call.
  Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
    logMessage("GetTypeAsync: " + asyncResult.value);
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("Action failed with error: " + asyncResult.error.message);
      logMessage("Action failed with error: " + asyncResult.error.message);
      return;
    }

    const bodyFormat = asyncResult.value;
    logMessage("Before prependAsync: " + contentToPrepend);
    Office.context.mailbox.item.body.prependAsync(contentToPrepend, { coercionType: bodyFormat }, (result) => {
      logMessage("Result Status: " + result.status);
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + result.error.message);
        logMessage("Action failed with error: " + result.error.message);
        return;
      }
      logMessage(contentToPrepend + " prepended to the body");
      console.log(`"${contentToPrepend}" prepended to the body.`);
    });
  });
}
