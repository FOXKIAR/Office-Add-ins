let snakeGameInterval;
let snake;
let food;
let direction;
let snakeScore = 0;
let gameRunning = false;

export function playGame() {
    Excel.run(async (context) => {
        const workSheet = context.workbook.worksheets.getItem("Sheet1");
        const gameRange = workSheet.getRange("A1:AZ25");
        gameRange.format.fill.color = "black";
        gameRange.format.columnWidth = 15;
        await context.sync();
        const playRange = workSheet.getRange("B2:AY24");
        playRange.format.fill.color = "white";
        playRange.values = "";
        await context.sync();

        startGame();
    });
}

function startGame() {
    // 初始化蛇（长度为3，水平放置）
    snake = [
        {x: 10, y: 12},
        {x: 11, y: 12},
        {x: 12, y: 12}
    ];

    // 初始方向向左
    direction = "LEFT";
    snakeScore = 0;
    gameRunning = true;

    // 生成第一个食物
    generateFood();

    // 添加键盘事件监听
    document.addEventListener("keydown", handleSnakeKeyDown);

    // 清除之前的游戏循环
    if (snakeGameInterval) clearInterval(snakeGameInterval);

    // 开始游戏循环
    snakeGameInterval = setInterval(() => {
        if (gameRunning) {
            moveSnake();
            renderSnake();
        }
    }, 200);
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
    if (!snake || !food) return;

    Excel.run(async (context) => {
        const workSheet = context.workbook.worksheets.getItem("Sheet1");
        const playRange = workSheet.getRange("B2:AY24");

        // 清除整个游戏区域
        playRange.format.fill.color = "white";

        // 绘制蛇
        snake.forEach((segment, index) => {
            if (segment.y >= 0 && segment.y < 23 && segment.x >= 0 && segment.x < 50) {
                const cell = playRange.getCell(segment.y, segment.x);
                cell.format.fill.color = index === 0 ? "green" : "lightgreen"; // 蛇头绿色，蛇身浅绿色
            }
        });

        // 绘制食物
        if (food.y >= 0 && food.y < 23 && food.x >= 0 && food.x < 50) {
            playRange.getCell(food.y, food.x).format.fill.color = "red";
        }

        // 更新分数
        const scoreCell = workSheet.getRange("AZ2");
        scoreCell.values = `得分: ${snakeScore}`;

        await context.sync();
    });
}

function gameOver() {
    gameRunning = false;
    clearInterval(snakeGameInterval);
    document.removeEventListener("keydown", handleSnakeKeyDown);

    Excel.run(async (context) => {
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
    });
}