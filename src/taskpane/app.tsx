// App.tsx
import * as React from "react";
//import { PrimaryButton, Checkbox, Stack, Text, TextField } from "@fluentui/react";
//import { Office } from "office-js";

import { PrimaryButton, Checkbox, Stack, Text, TextField, Pivot, PivotItem, Label } from "@fluentui/react";
import { initializeIcons } from "@fluentui/react/lib/Icons";

initializeIcons();

export const App = () => {
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [logMessages, setLogMessages] = React.useState<string[]>([]);
  const [nestedSelections, setNestedSelections] = React.useState<Record<string, string[]>>({});
  const checkboxOptions = ["Product Review", "Movies and Television", "Science and Technology"]; // You can customize this
  const [expandableFields, setExpandableFields] = React.useState<Record<string, string>>({});
  const nestedOptionsMap: Record<string, string[]> = {
    "Product Review": ["Consumer Electronics", "Laptop Computers", "Automotive Accessories", "Athletic Wear"],
    "Movies and Television": ["Science Fiction", "Romantic Comedies", "US Westerns", "Political Satire"],
    "Science and Technology": ["Artificial Intelligence", "Quantum Computing", "Space Exploration and Satellites", "Physics and Theororetical Mathematics"],
  };

  React.useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        runOutlook();
      }
    });
  }, []);

  const getFormattedDateTime = (): string => {
    const now = new Date();
    return now.toLocaleString("en-US", {
      weekday: "short",
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  const logMessage = (message: string) => {
    setLogMessages((prev) => [...prev, message]);
  };

  const handleMainCheckboxChange = (option: string, checked?: boolean) => {
    setSelectedOptions((prev) =>
      checked ? [...prev, option] : prev.filter((val) => val !== option)
    );
    if (!checked) {
      setNestedSelections((prev) => {
        const updated = { ...prev };
        delete updated[option];
        return updated;
      });
    }
  };

  const handleNestedCheckboxChange = (main: string, sub: string, checked?: boolean) => {
    setNestedSelections((prev) => {
      const current = prev[main] || [];
      const updated = checked
        ? [...current, sub]
        : current.filter((val) => val !== sub);
      return { ...prev, [main]: updated };
    });
  };

  const handleFieldChange = (option: string, value?: string) => {
    setExpandableFields((prev) => ({
      ...prev,
      [option]: value || "",
    }));
  };

  const runOutlook = () => {
    const item = Office.context.mailbox.item;
    logMessage(`Subject: ${item.subject}`);
  };

  const prependText = () => {
    if (selectedOptions.length === 0) {
      logMessage("No options selected.");
      return;
    }
    /*
  const contentToPrepend = selectedOptions
      .map((option) => {
        const nested = nestedSelections[option] || [];
        const nestedHtml = nested.length
          ? `<ul style="list-style-position: inside; padding-left: 0; margin: 10px 0 0 0;">${nested.map((item) => `<li style="text-align: center;">${item}</li>`).join("")}</ul>`
          : "";

        return `
          <div style="margin-bottom: 15px;">
            <div style="font-weight: bold; text-align: center; font-size: 18px;">${option}</div>
            ${nestedHtml}
          </div>`;
      })
      .join("");
  */
    const contentToPrepend = '<div style="text-align: center; font-weight: bold; margin-bottom: 10px;">' + selectedOptions
      .map((option) => {
        const nested = nestedSelections[option] || [];
        return [option, ...nested].join(" / ");
      })
      .join(" / ")+
      '</div>';

    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    logMessage(`UserEmail: ${userEmail}`);
    const userName = Office.context.mailbox.userProfile.displayName;
    logMessage(`UserName: ${userName}`);
    const generateSignatureBlock = (name: string, email: string): string => {
      const timestamp = getFormattedDateTime();
      return `
        <div style="margin-top: 20px; font-family: Arial, sans-serif; font-size: 12px; color: #444;">
          <div style="font-weight: bold;">${name}</div>
          <div>${email}</div>
          <div>MyCompany, Department Name</div>
          <div>📞 (123) 456-7890 | 🌐 www.mycompany.com</div>
          <div style="margin-top: 10px; font-style: italic;">Sent on ${timestamp}</div>
        </div>
      `;
    };
  const signatureHtml = generateSignatureBlock(userName, userEmail);

    const finalHtml = `
      ${contentToPrepend}
      <br/><br/>
      ${signatureHtml}
    `;

    Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        logMessage(`Error: ${asyncResult.error.message}`);
        return;
      }

      const bodyFormat = asyncResult.value;

    Office.context.mailbox.item.body.prependAsync(
        //contentToPrepend,
        finalHtml,
        { coercionType: bodyFormat },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            logMessage(`Prepend failed: ${result.error.message}`);
          } else {
            logMessage("Content prepended successfully.");
          }
        }
      );
    });
  };

  const wrapEmailBodyWithSelections = () => {
    if (selectedOptions.length === 0) {
      logMessage("No options selected.");
      return;
    }
  
    // Build the shared header/footer content
    const selectionSummary = selectedOptions
      .map((option) => {
        const nested = nestedSelections[option] || [];
        return [option, ...nested].join(" / ");
      })
      .join(" / ");
  
   
    const styledHtml = `
  <table width="100%" cellpadding="0" cellspacing="0" style="margin: 15px 0;">
    <tr>
      <td style="border-top: 1px solid #ccc; border-bottom: 1px solid #ccc; padding: 10px; font-weight: bold; text-align: center; font-family: Arial, sans-serif; font-size: 14px;">
        Selected Topics: ${selectionSummary}
      </td>
    </tr>
  </table>
`;

  
    // Get current body
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (getResult) => {
      if (getResult.status === Office.AsyncResultStatus.Succeeded) {
        const currentBody = getResult.value;
  
        // Wrap the email body
        const wrappedBody = `
          ${styledHtml}
          ${currentBody}
          ${styledHtml}
        `;
  
        // Set the updated body
        Office.context.mailbox.item.body.setAsync(
          wrappedBody,
          { coercionType: Office.CoercionType.Html },
          (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              logMessage("Email wrapped with header and footer successfully.");
            } else {
              logMessage(`SetAsync failed: ${setResult.error.message}`);
            }
          }
        );
      } else {
        logMessage(`GetAsync failed: ${getResult.error.message}`);
      }
    });
  };
  

  return (
    <Stack tokens={{ childrenGap: 15, padding: 20 }}>
      <Text variant="xLarge">Outlook Add-in Panel</Text>

      <Pivot>
        <PivotItem headerText="Options">
          <Stack tokens={{ childrenGap: 10 }}>
            {checkboxOptions.map((opt) => (
              <div key={opt}>
                <Checkbox
                  label={opt}
                  checked={selectedOptions.includes(opt)}
                  onChange={(_, checked) => handleMainCheckboxChange(opt, checked)}
                />
                {selectedOptions.includes(opt) &&
                  nestedOptionsMap[opt]?.map((sub) => (
                    <Checkbox
                      key={sub}
                      label={`↳ ${sub}`}
                      styles={{ root: { marginLeft: 20 } }}
                      checked={(nestedSelections[opt] || []).includes(sub)}
                      onChange={(_, checked) => handleNestedCheckboxChange(opt, sub, checked)}
                    />
                ))}
              </div>
            ))}
          </Stack>
        </PivotItem>

        <PivotItem headerText="Logs">
          <Stack tokens={{ childrenGap: 5 }}>
            {logMessages.length === 0 ? (
              <Text>No logs yet.</Text>
            ) : (
              logMessages.map((msg, idx) => <Text key={idx}>{msg}</Text>)
            )}
          </Stack>
        </PivotItem>
      </Pivot>

      <PrimaryButton text="Prepend to Email Body" onClick={prependText} />
      <PrimaryButton text="Wrap Email Body" onClick={wrapEmailBodyWithSelections} />
    </Stack>
  );
};
