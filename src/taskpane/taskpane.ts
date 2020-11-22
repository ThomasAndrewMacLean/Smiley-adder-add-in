/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
 
const smileys = ["😀","😆","😅","🙂","😄","😜","🙃"];

Office.onReady(info => {
  console.log("Start 🍾")
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = run;
    document.getElementById("run-remove").onclick = runRemove;
  }
});

export async function run() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, asyncResult => {
    if (asyncResult.status.toString() === "succeeded") {
      const lines = asyncResult.value.split("</div>");
      const newLines = lines.map(line => {
        // strip away html tags and trim to check if the line has text
        if (line.replace(/<.+?>/g, "").trim() !== "") {
          // regex to check if we find a smiley in this line
          const hasNoSmiley =
            line.search(
              /(\u00a9|\u00ae|[\u2000-\u3300]|\ud83c[\ud000-\udfff]|\ud83d[\ud000-\udfff]|\ud83e[\ud000-\udfff])/
            ) === -1;

          if (hasNoSmiley) {
            // add random smiley to the end of the line
            line += ` ${smileys[Math.floor(Math.random() * smileys.length)]}`;
          }
        }

        return line;
      });

      Office.context.mailbox.item.body.setAsync(newLines.join("</div>"), { coercionType: Office.CoercionType.Html });
    }
  });
}



export async function runRemove() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, asyncResult => {
    if (asyncResult.status.toString() === "succeeded") {
      const lines = asyncResult.value ;
    
      Office.context.mailbox.item.body.setAsync(lines.replace(/(\u00a9|\u00ae|[\u2000-\u3300]|\ud83c[\ud000-\udfff]|\ud83d[\ud000-\udfff]|\ud83e[\ud000-\udfff])/g,''), { coercionType: Office.CoercionType.Html });
    }
  });
}
