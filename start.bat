echo "请确保该设备已安装 node.js 并将 node.js 添加到设备 PATH 环境变量中。"
set /p isOk="准备好了吗？(Y/n)"
if %isOk%==N || %isOk%==n exit

echo "请选择要生成的任务窗格加载项(Word、PowerPoint、Excel)"
set /p select="w、p、e"
