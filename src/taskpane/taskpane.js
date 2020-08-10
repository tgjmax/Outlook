/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        // document.getElementById("sideload-msg").style.display = "none";
        // document.getElementById("app-body").style.display = "flex";
        // document.getElementById("run").onclick = run;
        // document.getElementById("get-master-categories").onclick = getMasterCategories;
        document.getElementById("add-master-categories").onclick = addMasterCategories;
        document.getElementById("remove-master-categories").onclick = removeMasterCategories;
        document.getElementById("remove-builtin-categories").onclick = removeBuiltinCategories
    }
});

// function getMasterCategories() {
//   Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
//       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//           var categories = asyncResult.value;
//           if (categories && categories.length > 0) {
//               console.log("Master categories:");
//               console.log(JSON.stringify(categories));
//           } else {
//               console.log("There are no categories in the master list.");
//           }
//       } else {
//           console.error(asyncResult.error);
//       }
//   });
// }

function addMasterCategories() {
    Swal.fire({
        title: 'Are you sure?',
        html:
            'New categories <b>Bot Mails, Non-Bot Mails</b>, ' +
            'will be added',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#28a745',
        cancelButtonColor: '#d33',
        confirmButtonText: 'Yes, add it!',
        customClass: {
            popup: 'swal2-popup2',
        }
    }).then((result) => {
        if (result.value) {
            var masterCategoriesToAdd = [
                {
                    displayName: "Non-Bot Mails",
                    color: Office.MailboxEnums.CategoryColor.Preset3
                },
                {
                    displayName: "Bot Mails",
                    color: Office.MailboxEnums.CategoryColor.Preset4
                }
            ];

            Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Successfully added categories to master list");
                } else {
                    console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
                }
            });
            Swal.fire({
                icon: 'success',
                width: 250,
                title: 'Added custom categories',
                showConfirmButton: false,
                timer: 1000,
                customClass: {
                    popup: 'swal2-popup2',
                }
            })
        }
    })
}

function removeMasterCategories() {
    Swal.fire({
        title: 'Are you sure?',
        html:
            'Categories <b>Bot Mails, Non-Bot Mails</b>, ' +
            'will be deleted',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#28a745',
        cancelButtonColor: '#d33',
        focusCancel: true,
        confirmButtonText: 'Delete',
        reverseButtons: true,
        customClass: {
            popup: 'swal2-popup2',
        }
    }).then((result) => {
        if (result.value) {
            var masterCategoriesToRemove = ["Non-Bot Mails", "Bot Mails"];

            Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Successfully removed categories from master list");
                } else {
                    console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
                }
            });
            Swal.fire({
                icon: 'success',
                width: 250,
                title: 'Deleted custom categories',
                showConfirmButton: false,
                timer: 1000,
                customClass: {
                    popup: 'swal2-popup2',
                }
            })
        }
    })

}

function removeBuiltinCategories() {
        Swal.fire({
        title: 'Are you sure?',
        text:'All built-in categories will be deleted',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#28a745',
        cancelButtonColor: '#d33',
        focusCancel: true,
        confirmButtonText: 'Delete',
        reverseButtons: true,
        customClass: {
            popup: 'swal2-popup2',
        }
    }).then((result) => {
        if (result.value) {
            var masterCategoriesToRemove = ["Green category", "Blue category", "Orange category", "Purple category", "Red category", "Yellow category"];

            Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Successfully removed categories from master list");
                } else {
                    console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
                }
            });
        
            Swal.fire({
                icon: 'success',
                width: 250,
                title: 'Deleted built-in categories',
                showConfirmButton: false,
                timer: 1000,
                customClass: {
                    popup: 'swal2-popup2',
                }
            })
        }
    })
}

// export async function run() {

// }
