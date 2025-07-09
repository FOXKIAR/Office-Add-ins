/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("rand").onclick = random;
    document.getElementById("default").onclick = none;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load("address");

      // Update the fill color.
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function random() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      let blue = Math.floor(Math.random() * 256);
      let red = Math.floor(Math.random() * 256);
      let green = Math.floor(Math.random() * 256);

      function toHex(c) {
        return c.toString(16).padStart(2, '0');
      }
      range.format.fill.color = `#${toHex(red)}${toHex(green)}${toHex(blue)}`;

      blue = Math.floor(Math.random() * 256);
      red = Math.floor(Math.random() * 256);
      green = Math.floor(Math.random() * 256);

      range.format.font.color = `#${toHex(red)}${toHex(green)}${toHex(blue)}`;
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function none() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");

      range.format.fill.color = '#FFFFFF';
      range.format.font.color = '#000000';
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}