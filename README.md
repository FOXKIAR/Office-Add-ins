#生成首个 Word 任务窗格加载项  
本文将逐步介绍如何在Windows环境中生成 Word 任务窗格加载项。  
## 先决条件  
1.安装 [Node.js](https://nodejs.org/)（最新LTS 版本）。  
2.在powshell中执行 `npm install -g yo generator-office`并等待执行完成。


## 使用说明
###1.通过下载本项目的方式运行：
1.1将下载下来的文件解压到非中文路径中，执行".\Office-Add-ins\word-add-in\start.bat"  
询问是否使用Microsoft Edge WebView 请选择n,如下图所示：
![](https://img.picgo.net/2024/05/02/-2024-05-02-1827257becff5b50c679eb.png)

询问是否安装CA证书 请选择 是。

1.2如果以上正常的话应该会打开word，并在窗口右侧弹出加载项任务窗格。 


###2.通过微软教程手动生成，完整教程见本文末
2.1在文件资源管理器的任意非中文路径中，右键打开powshell输入`yo office` 回车执行  
  如果遇到不能执行ps代码的报错，请参阅[文档链接](https://learn.microsoft.com/zh-cn/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.4)  
2.2出现提示时，请提供以下信息以创建加载项项目。

-  选择项目类型：Office Add-in Task Pane project
-  选择脚本类型：JavaScript
-  要为外接程序命名什么名称？My Office Add-in
-  你希望支持哪个 Office 客户端应用程序？Word   

2.3执行
进入到2.2所输入的外接程序命名名称文件夹内，右键打开powshell输入`npm start` 回车执行即可，详见本项目1.1start.bat之后。


微软教程：[https://learn.microsoft.com/zh-cn/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator](https://learn.microsoft.com/zh-cn/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator)



