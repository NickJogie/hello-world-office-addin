import * as React from "react";
import {
  PrimaryButton,
  Checkbox,
  Stack,
  Text,
  Pivot,
  PivotItem,
} from "@fluentui/react";
import { initializeIcons } from "@fluentui/react/lib/Icons";

initializeIcons();

const checkboxOptions = [
  "Product Review",
  "Movies and Television",
  "Science and Technology",
];

const nestedOptionsMap: Record<string, string[]> = {
  "Product Review": [
    "Consumer Electronics",
    "Laptop Computers",
    "Automotive Accessories",
    "Athletic Wear",
  ],
  "Movies and Television": [
    "Science Fiction",
    "Romantic Comedies",
    "US Westerns",
    "Political Satire",
  ],
  "Science and Technology": [
    "Artificial Intelligence",
    "Quantum Computing",
    "Space Exploration and Satellites",
    "Physics and Theoretical Mathematics",
  ],
};

const ORIGIN1 = "https://localhost:3000"; // update as needed

const App = () => {
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [nestedSelections, setNestedSelections] = React.useState<
    Record<string, string[]>
  >({});
  const [logMessages, setLogMessages] = React.useState<string[]>([]);

  const logMessage = (msg: string) =>
    setLogMessages((prev) => [...prev, msg]);

  // postMessage listener
  React.useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.origin !== ORIGIN1) return;

      const { type, payload } = event.data;
      if (type === "SET_SELECTIONS") {
        setSelectedOptions(payload || []);
        logMessage("Selections set via message from origin1.");
      }
    };

    window.addEventListener("message", handleMessage);
    return () => window.removeEventListener("message", handleMessage);
  }, []);

  // Send selections to parent
  const notifyParent = () => {
    if (window.parent) {
      window.parent.postMessage(
        {
          type: "SELECTIONS_UPDATED",
          payload: selectedOptions, nestedSelections,
        },
        ORIGIN1
      );
      logMessage("Sent selection update to origin1.");
    }
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

  const handleNestedCheckboxChange = (
    main: string,
    sub: string,
    checked?: boolean
  ) => {
    setNestedSelections((prev) => {
      const current = prev[main] || [];
      const updated = checked
        ? [...current, sub]
        : current.filter((val) => val !== sub);
      return { ...prev, [main]: updated };
    });
  };

  return (
    <Stack tokens={{ childrenGap: 15, padding: 20 }}>
      <Text variant="xLarge">Cross-Origin Form (origin2)</Text>

      <Pivot>
        <PivotItem headerText="Options">
          <Stack tokens={{ childrenGap: 10 }}>
            {checkboxOptions.map((opt) => (
              <div key={opt}>
                <Checkbox
                  label={opt}
                  checked={selectedOptions.includes(opt)}
                  onChange={(_, checked) =>
                    handleMainCheckboxChange(opt, checked)
                  }
                />
                {selectedOptions.includes(opt) &&
                  nestedOptionsMap[opt]?.map((sub) => (
                    <Checkbox
                      key={sub}
                      label={`↳ ${sub}`}
                      styles={{ root: { marginLeft: 20 } }}
                      checked={(nestedSelections[opt] || []).includes(sub)}
                      onChange={(_, checked) =>
                        handleNestedCheckboxChange(opt, sub, checked)
                      }
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

      <PrimaryButton text="Send Selection to Parent" onClick={notifyParent} />
    </Stack>
  );
};

export default App;
