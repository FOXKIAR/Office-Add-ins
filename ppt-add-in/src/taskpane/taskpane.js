/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { base64Image } from "../../base64Image";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("test").onclick = test;
    document.getElementById("flower").onclick = flower
  }
});

export async function run() {
  const options = { coercionType: Office.CoercionType.Text };

  Office.context.document.setSelectedDataAsync("Hello World!", options);
}

export async function test() {
  const options = { coercionType: Office.CoercionType.Image };
  // 直接插入base64图片
  Office.context.document.setSelectedDataAsync(base64Image, options);
}

export async function flower() {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const slideProperty = {
      height: 540,
      width: 960
    }
    
    const petalLeftOption = {
      left: slideProperty.width / 2 - 168-25,
      top: (slideProperty.height - 168) / 2
    }

    const petalTopOption = {
      left: (slideProperty.width - 168) / 2 ,
      top: slideProperty.height / 2 - 168-25
    }

    const petalRightOption = {
      left: petalLeftOption.left + 168+50,
      top: petalLeftOption.top
    }

    const petalBottomOption = {
      left: petalTopOption.left,
      top: petalTopOption.top + 168+50
    }

    for (const petalOption of [petalLeftOption, petalTopOption, petalRightOption, petalBottomOption]) {
      petalOption.height = 168; petalOption.width = 168;
      shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, petalOption)
    }
    
    const pistilOption = {
      left: (slideProperty.width - 100) / 2,
      top: (slideProperty.height - 100) / 2,
      height: 100,
      width: 100
    }
    // 花蕊
    const pistil = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, pistilOption);
    pistil.fill.foregroundColor = "#FFD966"

    await context.sync();
  });
}