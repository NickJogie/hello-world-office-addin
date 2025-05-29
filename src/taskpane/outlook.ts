/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";

  }
});

window.addEventListener("DOMContentLoaded", () => {
  const iframe = document.createElement("iframe");
  iframe.id = "formFrame";
  iframe.src = "https://localhost:4000";
  iframe.style.width = "100%";
  iframe.style.height = "600px";
  iframe.style.border = "none";
  document.getElementById("iframe-container")?.appendChild(iframe);
});

window.addEventListener("message", (event) => {
  if (event.origin !== "https://localhost:4000") return;

  const { type, payload, nestedSelections } = event.data;

  if (type === "SELECTIONS_UPDATED") {
    const output = `
      <strong>Main Selections:</strong> ${JSON.stringify(payload)}<br/>
      <strong>Nested Selections:</strong> ${JSON.stringify(nestedSelections)}
    `;
    document.getElementById("response").innerHTML = output;
    console.log("Selections from iframe:", { payload, nestedSelections });

    prependSelectionsToEmail(payload, nestedSelections);
  }
});

function requestIframeSelections() {
  const iframe = document.getElementById("formFrame") as HTMLIFrameElement;
  if (iframe?.contentWindow) {
    iframe.contentWindow.postMessage({ type: "REQUEST_SELECTIONS" }, "https://localhost:4000");
  }
}

function prependSelectionsToEmail(
  selectedOptions: string[],
  nestedSelections: Record<string, string[]>
) {
  // Combine selections and nested selections into a flat string
  const flatSelections = selectedOptions
    .map(option => {
      const nested = nestedSelections[option] || [];
      return [option, ...nested].join(" / ");
    })
    .join(" / ");

  // Create a single-line banner block
  const bannerHtml = `
    <div style="background: #e6f2ff; border-left: 6px solid #0078d4; padding: 16px; margin-bottom: 20px; font-family: Arial, sans-serif;">
      <div style="font-size: 14px; color: #333;">
        ${flatSelections}
      </div>
    </div>
  `;

  // Get current body format (HTML or text)
  Office.context.mailbox.item.body.getTypeAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to get body type:", result.error.message);
      return;
    }

    const bodyFormat = result.value;

    // Prepend the banner HTML
    Office.context.mailbox.item.body.prependAsync(
      bannerHtml,
      { coercionType: bodyFormat },
      (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to prepend content:", res.error.message);
        } else {
          console.log("Selections inserted into the email body.");
        }
      }
    );
  });
}

