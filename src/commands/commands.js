/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.warn("EEEEEEEEEEmmmmmmmmmRRRRRRRRReeeeeeeee")
});

let lastAttachmentCount = 0;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Polling yöntemiyle attachment var mı kontrol ederiz
    setInterval(() => {
      const item = Office.context.mailbox.item;

      if (item && item.attachments) {
        const currentCount = item.attachments.length;

        if (currentCount > lastAttachmentCount) {
          alert("⚠️ Dosya eklendi!");
          lastAttachmentCount = currentCount;
        } else if (currentCount < lastAttachmentCount) {
          lastAttachmentCount = currentCount;
        }
      }
    }, 1000); // her 1 saniyede bir kontrol
  }
});


/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
