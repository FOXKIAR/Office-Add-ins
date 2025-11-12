/* global Excel, OfficeExtension */

let snakeGameInterval;
let snake;
let food;
let direction;
let snakeScore = 0;
let gameRunning = false;

// 错误处理函数
function handleExcelError(operation, error) {
    console.error(`${operation} 失败:`, error);
    if (error instanceof OfficeExtension.Error) {
        console.error('Office Extension 错误:', error.debugInfo);
    }
    gameRunning = false;
}

export function playGame() {
    console.log("🐍 贪吃蛇游戏：开始初始化");

    Excel.run(async (context) => {
        try {
            console.log("🐍 贪吃蛇游戏：Excel.run() 成功启动");

            // 确保工作表存在
            let workSheet;
            try {
                console.log("🐍 贪吃蛇游戏：尝试获取工作表 Sheet1");
                workSheet = context.workbook.worksheets.getItem("Sheet1");
                workSheet.load("name");
                await context.sync();
                console.log("🐍 贪吃蛇游戏：成功获取工作表:", workSheet.name);
            } catch (error) {
                console.log("🐍 贪吃蛇游戏：工作表 Sheet1 不存在，创建新工作表");
                workSheet = context.workbook.worksheets.add("Sheet1");
                await context.sync();
                console.log("🐍 贪吃蛇游戏：创建工作表成功");
            }

            console.log("🐍 贪吃蛇游戏：开始设置游戏区域");
            const gameRange = workSheet.getRange("A1:AZ25");
            gameRange.format.fill.color = "black";
            gameRange.format.columnWidth = 15;
            await context.sync();
            console.log("🐍 贪吃蛇游戏：游戏边框设置完成");

            const playRange = workSheet.getRange("B2:AY24");
            playRange.format.fill.color = "white";
            playRange.values = "";
            await context.sync();
            console.log("🐍 贪吃蛇游戏：游戏区域初始化完成");

            console.log("🐍 贪吃蛇游戏：调用 startGame()");
            startGame();
            console.log("🐍 贪吃蛇游戏：初始化完成！");

        } catch (error) {
            console.error("🐍 贪吃蛇游戏：初始化失败", error);
            handleExcelError("初始化贪吃蛇游戏", error);
        }
    }).catch(error => {
        console.error("🐍 贪吃蛇游戏：Excel.run() 失败", error);
        handleExcelError("Excel操作", error);
    });
}

function startGame() {
    console.log("🐍 贪吃蛇游戏：startGame() 被调用");

    // 初始化蛇（长度为3，水平放置）
    snake = [
        {x: 10, y: 12},
        {x: 11, y: 12},
        {x: 12, y: 12}
    ];
    console.log("🐍 贪吃蛇游戏：蛇初始化完成", snake);

    // 初始方向向左
    direction = "LEFT";
    snakeScore = 0;
    gameRunning = true;
    console.log("🐍 贪吃蛇游戏：游戏状态初始化完成");

    // 生成第一个食物
    generateFood();
    console.log("🐍 贪吃蛇游戏：第一个食物生成完成", food);

    // 添加键盘事件监听
    console.log("🐍 贪吃蛇游戏：添加键盘事件监听器");
    document.addEventListener("keydown", handleSnakeKeyDown);
    console.log("🐍 贪吃蛇游戏：键盘事件监听器添加完成");

    // 清除之前的游戏循环
    if (snakeGameInterval) {
        console.log("🐍 贪吃蛇游戏：清除之前的游戏循环");
        clearInterval(snakeGameInterval);
    }

    // 开始游戏循环
    console.log("🐍 贪吃蛇游戏：开始游戏循环");
    snakeGameInterval = setInterval(() => {
        if (gameRunning) {
            console.log("🐍 贪吃蛇游戏：游戏循环运行中...");
            moveSnake();
            renderSnake();
        } else {
            console.log("🐍 贪吃蛇游戏：游戏已暂停，跳过循环");
        }
    }, 200);
    console.log("🐍 贪吃蛇游戏：游戏循环已启动，间隔200ms");
}

function generateFood() {
    const maxX = 49; // B到AY是50列，索引0-49
    const maxY = 22; // 2到24是23行，索引0-22

    let newFood;
    let foodOnSnake;

    do {
        foodOnSnake = false;
        newFood = {
            x: Math.floor(Math.random() * maxX),
            y: Math.floor(Math.random() * maxY)
        };

        // 检查食物是否生成在蛇身上
        for (const segment of snake) {
            if (segment.x === newFood.x && segment.y === newFood.y) {
                foodOnSnake = true;
                break;
            }
        }
    } while (foodOnSnake);

    food = newFood;
}

function moveSnake() {
    // 创建新的蛇头
    const head = {...snake[0]};

    switch(direction) {
        case "UP":
            head.y -= 1;
            break;
        case "DOWN":
            head.y += 1;
            break;
        case "LEFT":
            head.x -= 1;
            break;
        case "RIGHT":
            head.x += 1;
            break;
    }

    // 检查碰撞
    if (checkCollision(head)) {
        gameOver();
        return;
    }

    // 将新头添加到蛇身
    snake.unshift(head);

    // 检查是否吃到食物
    if (head.x === food.x && head.y === food.y) {
        snakeScore += 10;
        generateFood();
    } else {
        // 如果没有吃到食物，移除蛇尾
        snake.pop();
    }
}

function checkCollision(head) {
    // 检查墙壁碰撞
    const maxX = 49;
    const maxY = 22;

    if (head.x < 0 || head.x > maxX || head.y < 0 || head.y > maxY) {
        return true;
    }

    // 检查自身碰撞
    for (let i = 0; i < snake.length; i++) {
        if (i !== 0 && head.x === snake[i].x && head.y === snake[i].y) {
            return true;
        }
    }

    return false;
}

function handleSnakeKeyDown(event) {
    if (!gameRunning) return;

    switch(event.code) {
        case "ArrowUp":
            if (direction !== "DOWN") direction = "UP";
            break;
        case "ArrowDown":
            if (direction !== "UP") direction = "DOWN";
            break;
        case "ArrowLeft":
            if (direction !== "RIGHT") direction = "LEFT";
            break;
        case "ArrowRight":
            if (direction !== "LEFT") direction = "RIGHT";
            break;
        case "Space":
            // 空格键暂停/继续
            gameRunning = !gameRunning;
            break;
    }
}

function renderSnake() {
    console.log("🐍 贪吃蛇游戏：renderSnake() 被调用");
    if (!snake || !food) {
        console.log("🐍 贪吃蛇游戏：蛇或食物不存在，跳过渲染");
        return;
    }

    console.log("🐍 贪吃蛇游戏：开始渲染，蛇长度:", snake.length, "食物位置:", food);

    Excel.run(async (context) => {
        try {
            console.log("🐍 贪吃蛇游戏：获取工作表和范围");
            const workSheet = context.workbook.worksheets.getItem("Sheet1");
            const playRange = workSheet.getRange("B2:AY24");

            console.log("🐍 贪吃蛇游戏：清除游戏区域");
            // 清除整个游戏区域
            playRange.format.fill.color = "white";

            console.log("🐍 贪吃蛇游戏：绘制蛇");
            // 绘制蛇
            snake.forEach((segment, index) => {
                if (segment.y >= 0 && segment.y < 23 && segment.x >= 0 && segment.x < 50) {
                    const cell = playRange.getCell(segment.y, segment.x);
                    cell.format.fill.color = index === 0 ? "green" : "lightgreen"; // 蛇头绿色，蛇身浅绿色
                }
            });

            console.log("🐍 贪吃蛇游戏：绘制食物");
            // 绘制食物
            if (food.y >= 0 && food.y < 23 && food.x >= 0 && food.x < 50) {
                playRange.getCell(food.y, food.x).format.fill.color = "red";
            }

            console.log("🐍 贪吃蛇游戏：更新分数显示");
            // 更新分数
            const scoreCell = workSheet.getRange("AZ2");
            scoreCell.values = `得分: ${snakeScore}`;

            console.log("🐍 贪吃蛇游戏：同步到Excel");
            await context.sync();
            console.log("🐍 贪吃蛇游戏：渲染完成！");

        } catch (error) {
            console.error("🐍 贪吃蛇游戏：渲染失败", error);
            if (error instanceof OfficeExtension.Error) {
                console.error('🐍 Office Extension 错误:', error.debugInfo);
            }
            // 渲染失败时停止游戏
            gameRunning = false;
        }
    }).catch(error => {
        console.error("🐍 贪吃蛇游戏：Excel渲染操作失败", error);
        gameRunning = false;
    });
}

function gameOver() {
    gameRunning = false;
    clearInterval(snakeGameInterval);
    document.removeEventListener("keydown", handleSnakeKeyDown);

    Excel.run(async (context) => {
        try {
            const workSheet = context.workbook.worksheets.getItem("Sheet1");
            const message = "GAME OVER!";
            const range = workSheet.getRange("B12:K13");
            range.format.fill.color = "black";
            range.format.font.color = "red";
            range.format.font.bold = true;

            // 清空区域
            range.values = "";

            // 显示游戏结束信息
            for (let i = 0; i < message.length; i++) {
                range.getCell(0, i).values = message[i];
            }

            // 显示最终分数
            const scoreMessage = `得分: ${snakeScore}`;
            const scoreRange = workSheet.getRange("B14:K14");
            scoreRange.format.fill.color = "black";
            scoreRange.format.font.color = "white";

            for (let i = 0; i < scoreMessage.length; i++) {
                scoreRange.getCell(0, i).values = scoreMessage[i];
            }

            await context.sync();
        } catch (error) {
            console.error("贪吃蛇游戏结束显示失败:", error);
            if (error instanceof OfficeExtension.Error) {
                console.error('Office Extension 错误:', error.debugInfo);
            }
        }
    }).catch(error => {
        console.error("Excel操作失败:", error);
    });
}