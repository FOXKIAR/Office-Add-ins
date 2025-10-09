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
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Text };

  Office.context.document.setSelectedDataAsync("Hello World!", options);
}

export async function test() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Image };
  // 直接插入base64图片
  Office.context.document.setSelectedDataAsync(base64Image, options);
}

export async function flower() {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    // 幻灯片尺寸
    //   height: 540,
    //   width: 960

    // 花瓣-左
    shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: 300,
      top: 220,
      height: 100,
      width: 150
    });

    // 旋转属性不生效，后续再修
    // // 花瓣-左上
    // const petal_left_top = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
    //   left: 330,
    //   top: 250,
    //   height: 100,
    //   width: 150
    // });
    
    
    // 花瓣-上
    shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: 430,
      top: 90,
      height: 150,
      width: 100
    });

    // 花瓣-右
    shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: 510,
      top: 220,
      height: 100,
      width: 150
    });

    // 花瓣-下
    shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: 430,
      top: 300,
      height: 150,
      width: 100
    });

    // 花蕊
    const pistil = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: 430,
      top: 220,
      height: 100,
      width: 100
    });
    pistil.fill.foregroundColor = "#FFD966"

    await context.sync();
  });
}