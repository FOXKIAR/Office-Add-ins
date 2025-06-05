# 生成首个 office 任务窗格加载项  
本文将逐步介绍如何在Windows环境中生成 (Word、PowerPoint、Excel) 任务窗格加载项。  
## 先决条件  
1.安装 [Node.js](https://nodejs.org/)（最新LTS 版本）。  
2.在powershell中执行 `npm install -g yo generator-office`并等待执行完成。
  如果遇到不能执行ps代码的报错，请参阅 [https:/go.microsoft.com/fwlink/?LinkID=135170](https:/go.microsoft.com/fwlink/?LinkID=135170)。


## 使用说明
### 1.通过下载本项目的方式运行：
1.1将下载下来的文件解压到非中文路径中，执行下列任意一项
+ ".\word-add-in\start.bat"
+ ".\excel-add-in\start.bat"
+ ".\ppt-add-in\start.bat"  

询问是否使用Microsoft Edge WebView 请选择y,如下图所示：
![](https://s21.ax1x.com/2025/02/05/pEeJ1QU.png)

询问是否安装CA证书 请选择 是
![](https://s21.ax1x.com/2025/02/05/pEeJQzT.png)

1.2如果以上正常的话应该会打开word，并在窗口右侧弹出加载项任务窗格。 


### 2.通过微软教程手动生成，完整教程见本文末
(前提是要先完成"先决条件"内的操作)  
2.1在文件资源管理器的任意非中文路径中，右键打开PowerShell输入`yo office` 回车执行  
  如果遇到不能执行ps代码的报错，请参阅[文档链接](https:/go.microsoft.com/fwlink/?LinkID=135170)。

2.2出现提示时，请提供以下信息以创建加载项项目。

-  选择项目类型：Office Add-in Task Pane project
-  选择脚本类型：JavaScript
-  要为外接程序命名什么名称？word-add-in  
-  你希望支持哪个 Office 客户端应用程序？Word   

2.3执行
进入到2.2所输入的外接程序命名名称文件夹内，右键打开PowerShell输入`npm start` 回车执行即可，详见本项目1.1start.bat之后。

## 3.后续开发请参阅微软相关教程
[生成首个 Word 任务窗格加载项](https://learn.microsoft.com/zh-cn/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator)   
[创建 Word 任务窗格加载项](https://learn.microsoft.com/zh-cn/office/dev/add-ins/tutorials/word-tutorial)
