/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  console.log("OUTLOOOOOOOOOOOOK INFO", info);
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
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

export async function run() {
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
