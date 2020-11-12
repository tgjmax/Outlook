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
        var cat_list = document.getElementById('cat-list').getAttribute("data-cat");
        if(!cat_list){
            document.getElementById("add-master-categories").disabled = true;
            document.getElementById("remove-master-categories").disabled = true;
            document.getElementById("remove-builtin-categories").disabled = true;
        }

        document.getElementById("email").value = Office.context.mailbox.userProfile.emailAddress
        document.getElementById("get-master-categories").onclick = getMasterCategories;
        document.getElementById("add-master-categories").onclick = addMasterCategories;
        document.getElementById("remove-master-categories").onclick = removeMasterCategories;
        document.getElementById("remove-builtin-categories").onclick = removeBuiltinCategories
    }
});

function getMasterCategories() {
    $(document).ready(function() {
        $.ajax({
          url: "http://f367b06b76a6.ngrok.io/api/categories",
          type: "POST",
          data: {
              email: $('#email').val()
          },
        //   xhrFields: {
        //       withCredentials: true
        //     },
        //   crossDomain: true,
          success: function(data) {
            console.log(data['categories']);
            console.log(data['color']);

            $("#cat-list").attr("data-cat", data['categories']);
            $("#cat-list").attr("data-color", data['color']);

            document.getElementById("add-master-categories").disabled = false;
            document.getElementById("remove-master-categories").disabled = false;
            document.getElementById("remove-builtin-categories").disabled = false;
          },
          error: function(error) {
            console.log("in error");
            console.log(error);
          }
        });
      });
}

function addMasterCategories() {
    Swal.fire({
        title: 'Are you sure?',
        html: 'New categories <b>' + $("#cat-list").attr("data-cat") + '</b>, will be added',
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
            var catArr = $("#cat-list").attr("data-cat").split(',');
            var colorArr = $("#cat-list").attr("data-color").split(',');
            var masterCategoriesToAdd = [];
            
            for (var i = 0; i < catArr.length; i++) {
                let catCol = "Preset" + colorArr[i]
                masterCategoriesToAdd.push({
                    displayName: catArr[i],
                    color: Office.MailboxEnums.CategoryColor[catCol]
                })
            }

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
        html: 'Categories <b>' + $("#cat-list").attr("data-cat") + '</b>, will be deleted',
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
            var masterCategoriesToRemove = $("#cat-list").attr("data-cat").split(',');

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
