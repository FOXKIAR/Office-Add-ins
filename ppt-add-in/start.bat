CHCP 65001

@echo off

echo "尝试下载依赖项"
call npm install || (
	echo "请确保该设备已安装`node.js`或将`node.js`可执行文件目录添加到设备`PATH`环境变量中后重试。"
	goto END
) && (
	npm start
)
:END 
