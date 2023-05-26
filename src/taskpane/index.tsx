import { App } from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

initializeIcons();

let subject = "";
let attendees = [];

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component subject={subject} attendees={[attendees]} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  const item = Office.context.mailbox.item;
  subject = item.subject;
  attendees = item.requiredAttendees.concat(item.optionalAttendees).map((val) => {
    return {
      name: val.displayName,
      email: val.emailAddress,
    };
  });
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
