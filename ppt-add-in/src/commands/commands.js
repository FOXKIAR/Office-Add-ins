/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  // If needed, Office.js is ready to be called.
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("run").onclick = run;
  }
});

function run(){
  Office.context.document.setSelectedDataAsync("Hello World!");
}
