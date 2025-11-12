import { createBlock } from "./blocks";
import { Board } from "./board";

/* global Excel, OfficeExtension */

// Excel加载项游戏主逻辑
let board;
let currentBlock;
let gameInterval;
let keyDownHandler;
let score = 0;
let isGameRunning = false;

// 错误处理函数
function handleExcelError(operation, error) {
    console.error(`${operation} 失败:`, error);
    if (error instanceof OfficeExtension.Error) {
        console.error('Office Extension 错误:', error.debugInfo);
    }
    stopGame();
}

// 确保工作表存在
async function ensureWorksheetExists(context, sheetName) {
    try {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.load("name");
        await context.sync();
        return sheet;
    } catch (error) {
        // 如果工作表不存在，创建新工作表
        const newSheet = context.workbook.worksheets.add(sheetName);
        await context.sync();
        return newSheet;
    }
}

export function playGame() {
    // 如果游戏正在运行，先停止当前游戏
    stopGame();

    Excel.run(async (context) => {
        try {
            // 确保工作表存在
            let workSheet;
            try {
                workSheet = context.workbook.worksheets.getItem("Sheet1");
            } catch (error) {
                workSheet = context.workbook.worksheets.add("Sheet1");
            }

            const gameRange = workSheet.getRange("A1:L22");
            gameRange.format.fill.color = "black";
            gameRange.format.columnWidth = 15;
            await context.sync();
            const playRange = workSheet.getRange("B2:K21");
            playRange.format.fill.color = "white";
            await context.sync();
            board = new Board(10, 20);

            startGame();
        } catch (error) {
            handleExcelError("初始化游戏", error);
        }
    }).catch(error => {
        handleExcelError("Excel操作", error);
    });
}

function startGame() {
    // 清理之前的游戏状态
    stopGame();
    
    isGameRunning = true;
    score = 0;
    currentBlock = createBlock(board);
    
    // 创建命名的事件处理函数，以便后续移除
    keyDownHandler = (event) => {
        if (!isGameRunning) return;
        
        if (event.code === "ArrowUp") {
            blockMove("rotate");
        } else if (event.code === "ArrowLeft") {
            blockMove("left");
        } else if (event.code === "ArrowRight") {
            blockMove("right");
        } else if (event.code === "ArrowDown") {
            blockMove("down");
        }
    };
    
    document.addEventListener("keydown", keyDownHandler);
    
    if (gameInterval) clearInterval(gameInterval);
    gameInterval = setInterval(() => {
        if (!isGameRunning || !currentBlock) {
            stopGame();
            return;
        }
        
        if (!currentBlock.down()) {
            board.placeBlock(currentBlock);
            const linesCleared = board.clearLines();
            score += linesCleared * 100;
            currentBlock = createBlock(board);

            // 检查游戏结束
            if (!board.isValidPosition(currentBlock.space)) {
                stopGame();
                console.log(`游戏结束！得分: ${score}`);
            }
        }
        renderBoard();
    }, 500);
    
    renderBoard();
}

function stopGame() {
    isGameRunning = false;
    
    if (gameInterval) {
        clearInterval(gameInterval);
        gameInterval = null;
    }
    
    if (keyDownHandler) {
        document.removeEventListener("keydown", keyDownHandler);
        keyDownHandler = null;
    }
}

function handleKeyDown(event) {
    if (event.code === "ArrowUp")
        blockMove("rotate");
    if (event.code === "ArrowLeft")
        blockMove("left");
    if (event.code === "ArrowRight")
        blockMove("right");
    if (event.code === "ArrowDown")
        blockMove("down");
}

function blockMove(direction) {
    if (!board || !currentBlock || !isGameRunning) return;
    
    let moveSuccess = false;
    
    switch(direction) {
        case "rotate":
            moveSuccess = currentBlock.rotate();
            break;
        case "left":
            moveSuccess = currentBlock.left();
            break;
        case "right":
            moveSuccess = currentBlock.right();
            break;
        case "down":
            moveSuccess = currentBlock.down();
            break;
    }
    
    // 只有移动成功时才渲染
    if (moveSuccess) {
        renderBoard();
    }
}

function renderBoard() {
    if (!board || !currentBlock || !isGameRunning) return;

    Excel.run(async (context) => {
        try {
            const workSheet = context.workbook.worksheets.getItem("Sheet1");
            const playRange = workSheet.getRange("B2:K21");

            // 清除整个游戏区域
            playRange.format.fill.color = "white";

            // 绘制固定方块
            for (let y = 0; y < board.height; y++) {
                for (let x = 0; x < board.width; x++) {
                    if (board.grid[y][x]) {
                        playRange.getCell(y, x).format.fill.color = board.grid[y][x];
                    }
                }
            }

            // 绘制当前活动方块
            currentBlock.space.forEach(pos => {
                if (pos.y >= 0 && pos.y < board.height && pos.x >= 0 && pos.x < board.width) {
                    playRange.getCell(pos.y, pos.x).format.fill.color = currentBlock.color;
                }
            });

            // 更新分数
            const scoreCell = workSheet.getRange("M2");
            scoreCell.values = [[`得分: ${score}`]];

            await context.sync();
        } catch (error) {
            console.error("渲染游戏板失败:", error);
            if (error instanceof OfficeExtension.Error) {
                console.error('Office Extension 错误:', error.debugInfo);
            }
            // 渲染失败时停止游戏
            stopGame();
        }
    }).catch(error => {
        console.error("Excel渲染操作失败:", error);
        stopGame();
    });
}

function gameOver() {
    Excel.run(async (context) => {
        try {
            const message = "GAME OVER!";
            const range = context.workbook.worksheets.getItem("Sheet1").getRange("B2:K3");
            range.format.fill.color = "black";
            range.format.font.color = "red";
            for (let i = 0; i < message.length; i++) {
                range.getCell(0, i).values = message[i];
            }
            await context.sync();
        } catch (error) {
            console.error("游戏结束显示失败:", error);
            if (error instanceof OfficeExtension.Error) {
                console.error('Office Extension 错误:', error.debugInfo);
            }
        }
    }).catch(error => {
        console.error("Excel操作失败:", error);
    });
}