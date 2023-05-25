/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import axios from "axios";

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function action(event: Office.AddinCommands.Event) {
  try {
    let response = await axios("https://alphanumericadvancedkeyboardmapping.zak0749.repl.co");
    let data = response.data;

    let res = await axios.post("https://alphanumericadvancedkeyboardmapping.zak0749.repl.co/data");

    const message: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: data.person + " says " + data.message + " status is" + res.status,
      icon: "Icon.80x80",
      persistent: true,
    };

    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
  } catch (error) {
    const message: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: error,
      icon: "Icon.80x80",
      persistent: true,
    };

    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

    event.completed();
  }
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
