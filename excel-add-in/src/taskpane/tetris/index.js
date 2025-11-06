import { createBlock } from "./blocks";
import { Board } from "./board";

let board;
let currentBlock;
let gameInterval;
let keyDownHandler;
let score = 0;
let isGameRunning = false;

export function playGame() {
    // 如果游戏正在运行，先停止当前游戏
    stopGame();
    
    try {
        Excel.run(async (context) => {
            const workSheet = context.workbook.worksheets.getItem("Sheet1");
            const gameRange = workSheet.getRange("A1:L22");
            gameRange.format.fill.color = "black";
            gameRange.format.columnWidth = 15;
            await context.sync();
            const playRange = workSheet.getRange("B2:K21");
            playRange.format.fill.color = "white";
            await context.sync();
            board = new Board(10, 20);

            startGame();
        }).catch(error => {
            console.error("初始化游戏失败:", error);
        });
    } catch (error) {
        console.error("游戏启动错误:", error);
    }
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
    
    try {
        Excel.run(async (context) => {
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
        }).catch(error => {
            console.error("渲染游戏板失败:", error);
        });
    } catch (error) {
        console.error("渲染错误:", error);
    }
}
