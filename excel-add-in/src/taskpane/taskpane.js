/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { playGame as playTetris } from "./tetris";
import { playGame as playSnake } from "./snake";

/* global console, document, Excel, Office, OfficeExtension */

// 全局错误处理
window.addEventListener('error', (event) => {
  console.error('🚨 全局错误捕获:', event.error);
  console.error('🚨 错误信息:', event.message);
  console.error('🚨 错误文件:', event.filename);
  console.error('🚨 错误行号:', event.lineno);
  console.error('🚨 错误列号:', event.colno);
  console.error('🚨 错误堆栈:', event.error?.stack);

  // 显示用户友好的错误信息
  alert(`脚本错误: ${event.message}\n请查看控制台获取详细信息。`);
});

// 未处理的Promise拒绝
window.addEventListener('unhandledrejection', (event) => {
  console.error('🚨 未处理的Promise拒绝:', event.reason);
  console.error('🚨 Promise拒绝详情:', event);

  // 显示用户友好的错误信息
  alert(`Promise错误: ${event.reason}\n请查看控制台获取详细信息。`);
});

// Office.js 特定的错误处理
if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    console.log("🎯 Office.onReady() 被调用", info);
  }).catch((error) => {
    console.error('🚨 Office.onReady() 失败:', error);
    alert('Office加载失败，请确保在Excel环境中运行。');
  });
} else {
  console.error('🚨 Office.js 未加载');
  alert('Office.js 库未加载，请检查网络连接。');
}

Office.onReady((info) => {
  console.log("🎯 Office.onReady() 被调用", info);

  if (info.host === Office.HostType.Excel) {
    console.log("✅ 确认运行在Excel环境中");

    try {
      // 检查模块导入状态
      console.log("🎯 检查模块导入状态:");
      console.log("- playTetris:", typeof playTetris);
      console.log("- playSnake:", typeof playSnake);

      if (typeof playTetris !== 'function' || typeof playSnake !== 'function') {
        throw new Error("游戏模块导入失败");
      }

      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      console.log("🎯 绑定按钮事件...");

      // 绑定各个按钮事件
      const runBtn = document.getElementById("run");
      const randBtn = document.getElementById("rand");
      const tetrisBtn = document.getElementById("tetris");
      const snakeBtn = document.getElementById("snake");

      console.log("🎯 按钮元素获取结果:");
      console.log("- run按钮:", runBtn);
      console.log("- rand按钮:", randBtn);
      console.log("- tetris按钮:", tetrisBtn);
      console.log("- snake按钮:", snakeBtn);

      if (runBtn) {
        runBtn.onclick = run;
        console.log("✅ run按钮事件绑定成功");
      } else {
        console.error("❌ run按钮元素未找到");
      }

      if (randBtn) {
        randBtn.onclick = random;
        console.log("✅ rand按钮事件绑定成功");
      } else {
        console.error("❌ rand按钮元素未找到");
      }

      if (tetrisBtn) {
        tetrisBtn.onclick = tetris;
        console.log("✅ tetris按钮事件绑定成功");
      } else {
        console.error("❌ tetris按钮元素未找到");
      }

      if (snakeBtn) {
        snakeBtn.onclick = snake;
        console.log("✅ snake按钮事件绑定成功");
      } else {
        console.error("❌ snake按钮元素未找到");
      }

      // 绑定调试按钮
      const debugBtn = document.getElementById("debug");
      if (debugBtn) {
        debugBtn.onclick = debugTest;
        console.log("✅ debug按钮事件绑定成功");
      } else {
        console.error("❌ debug按钮元素未找到");
      }

      console.log("🎯 所有按钮事件绑定完成！");

      // 显示调试信息
      console.log("🎯 Office.js 版本:", Office.context.diagnostics?.version || "未知");
      console.log("🎯 Office 主机:", Office.context.diagnostics?.host || "未知");
      console.log("🎯 Office 平台:", Office.context.diagnostics?.platform || "未知");

    } catch (error) {
      console.error("🎯 按钮事件绑定失败:", error);
    }
  } else {
    console.warn("⚠️ 不在Excel环境中，当前环境:", info.host);
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

export function tetris() {
  console.log("🎯 俄罗斯方块按钮被点击！");
  try {
    playTetris();
    console.log("✅ 俄罗斯方块游戏启动成功");
  } catch (error) {
    console.error("❌ 俄罗斯方块游戏启动失败:", error);
  }
}

export function snake() {
  console.log("🎯 贪吃蛇按钮被点击！");
  try {
    playSnake();
    console.log("✅ 贪吃蛇游戏启动成功");
  } catch (error) {
    console.error("❌ 贪吃蛇游戏启动失败:", error);
  }
}

export function debugTest() {
  console.log("🐛 调试按钮被点击！");

  // 测试1: 检查Office是否可用
  console.log("🐛 测试1: Office对象状态", typeof Office);
  console.log("🐛 测试1: Excel对象状态", typeof Excel);
  console.log("🐛 测试1: OfficeExtension对象状态", typeof OfficeExtension);

  // 测试2: 检查文档元素和按钮
  console.log("🐛 测试2: 当前文档元素数量", document.querySelectorAll("*").length);
  console.log("🐛 测试2: run按钮元素", document.getElementById("run"));
  console.log("🐛 测试2: rand按钮元素", document.getElementById("rand"));
  console.log("🐛 测试2: tetris按钮元素", document.getElementById("tetris"));
  console.log("🐛 测试2: snake按钮元素", document.getElementById("snake"));
  console.log("🐛 测试2: debug按钮元素", document.getElementById("debug"));
  console.log("🐛 测试2: Office.js是否已加载完成", Office.context !== undefined);

  // 测试3: 简单的Excel操作测试
  try {
    Excel.run(async (context) => {
      console.log("🐛 测试3: Excel.run() 成功启动");
      const workSheet = context.workbook.worksheets.getItem("Sheet1");
      const testRange = workSheet.getRange("A1");
      testRange.values = [["调试测试成功！"]];
      testRange.format.fill.color = "yellow";
      await context.sync();
      console.log("🐛 测试3: Excel操作成功完成！");
    }).catch(error => {
      console.error("🐛 测试3: Excel操作失败:", error);
      console.error("🐛 测试3: 错误详情:", {
        name: error.name,
        message: error.message,
        code: error.code,
        stack: error.stack
      });
      if (error instanceof OfficeExtension.Error) {
        console.error("🐛 测试3: OfficeExtension错误详情:", error.debugInfo);
      }
    });
  } catch (error) {
    console.error("🐛 测试3: Excel.run() 失败:", error);
    console.error("🐛 测试3: 错误类型:", error.constructor.name);
  }

  // 测试4: 模块导入状态检查
  try {
    console.log("🐛 测试4: 检查游戏模块导入状态");
    console.log("🐛 测试4: playTetris函数:", typeof playTetris);
    console.log("🐛 测试4: playSnake函数:", typeof playSnake);
  } catch (error) {
    console.error("🐛 测试4: 模块检查失败:", error);
  }

  // 测试5: 弹出提示
  alert("调试测试完成！请查看控制台日志。\n如果看到错误，请复制错误信息给我。");
}