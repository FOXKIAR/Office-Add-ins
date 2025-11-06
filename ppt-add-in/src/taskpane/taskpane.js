/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { base64Image } from "../../base64Image";

/* global document, Office, PowerPoint */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = () => tryCatch(run);
    document.getElementById("test").onclick = () => tryCatch(test);
    document.getElementById("flower").onclick = () => tryCatch(flower);
  }
});

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

export async function run() {
  const options = { coercionType: Office.CoercionType.Text };

  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync("Hello World!", options, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve(result);
      }
    });
  });
}

export async function test() {
  const options = { coercionType: Office.CoercionType.Image };
  // 直接插入base64图片
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64Image, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve(result);
      }
    });
  });
}

export async function flower() {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.warn("演示文稿中没有幻灯片");
        return;
      }

      const shapes = slides.getItemAt(0).shapes;
      const slideProperty = {
        height: 540,
        width: 960
      };
      
      const petalLeftOption = {
        left: slideProperty.width / 2 - 168 - 25,
        top: (slideProperty.height - 168) / 2
      };

      const petalTopOption = {
        left: (slideProperty.width - 168) / 2,
        top: slideProperty.height / 2 - 168 - 25
      };

      const petalRightOption = {
        left: petalLeftOption.left + 168 + 50,
        top: petalLeftOption.top
      };

      const petalBottomOption = {
        left: petalTopOption.left,
        top: petalTopOption.top + 168 + 50
      };

      for (const petalOption of [petalLeftOption, petalTopOption, petalRightOption, petalBottomOption]) {
        petalOption.height = 168;
        petalOption.width = 168;
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, petalOption);
      }
      
      const pistilOption = {
        left: (slideProperty.width - 100) / 2,
        top: (slideProperty.height - 100) / 2,
        height: 100,
        width: 100
      };
      // 花蕊
      const pistil = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, pistilOption);
      pistil.fill.foregroundColor = "#FFD966";

      await context.sync();
    });
  } catch (error) {
    console.error("创建花朵形状失败:", error);
    throw error;
  }
}