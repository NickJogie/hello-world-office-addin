import * as React from 'react';
import * as ReactDOM from 'react-dom/client';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { PrimaryButton, Checkbox, TextField, Label } from '@fluentui/react';
import { App } from './app';

initializeIcons();  // Initialize Fluent UI Icons

const root = ReactDOM.createRoot(document.getElementById("root")!);
root.render(<App />);
