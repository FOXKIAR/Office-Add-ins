/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { logo } from "../../assets/logo-filled.png"

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("test").onclick = function () {Office.context.document.setSelectedDataAsync(" show test ", options)};
  }
});

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Matrix };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

export async function test() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Image };

  const reader = new FileReader();
  reader.readAsDataURL(logo);

  const logoBase64Str = reader.result.split(",")[1];

  await Office.context.document.setSelectedDataAsync(logoBase64Str, options);
}