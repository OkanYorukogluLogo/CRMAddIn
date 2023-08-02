/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

 
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.addEventListener("DOMContentLoaded", function() {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      //document.getElementById("run").onclick = run;
      document.getElementById("myButton").onclick = onButtonClick;
    });
  }
});

 
/* eslint-disable */
function onButtonClick() {
    console.log("sdfsd");
}
