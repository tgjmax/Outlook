/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("get-master-categories").onclick = getMasterCategories;
    document.getElementById("add-master-categories").onclick = addMasterCategories;
    document.getElementById("remove-master-categories").onclick = removeMasterCategories;
    document.getElementById("remove-builtin-categories").onclick = removeBuiltinCategories
  }
});

function getMasterCategories() {
  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          var categories = asyncResult.value;
          if (categories && categories.length > 0) {
              console.log("Master categories:");
              console.log(JSON.stringify(categories));
          } else {
              console.log("There are no categories in the master list.");
          }
      } else {
          console.error(asyncResult.error);
      }
  });
}

function addMasterCategories() {
  var masterCategoriesToAdd = [
      {
          displayName: "Urgent Mails",
          color: Office.MailboxEnums.CategoryColor.Preset0
      },
      {
          displayName: "Bot Mails",
          color: Office.MailboxEnums.CategoryColor.Preset1
      },
      {
          displayName: "Invoice Mails",
          color: Office.MailboxEnums.CategoryColor.Preset3
      },
      {
          displayName: "General Mails",
          color: Office.MailboxEnums.CategoryColor.Preset4
      },
  ];

  Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Successfully added categories to master list");
      } else {
          console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
      }
  });
}

function removeMasterCategories() {
  var masterCategoriesToRemove = ["Urgent Mails", "Bot Mails", "Invoice Mails", "General Mails"];

  Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Successfully removed categories from master list");
      } else {
          console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
      }
  });
}

function removeBuiltinCategories() {
    var masterCategoriesToRemove = ["Green category", "Blue category", "Orange category", "Purple category", "Red category", "Yellow category"];
  
    Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Successfully removed categories from master list");
        } else {
            console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
        }
    });
  }

export async function run() {

}
