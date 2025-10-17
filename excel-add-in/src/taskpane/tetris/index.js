import { createBlock } from "./blocks";
import { Board } from "./board";

// Excel加载项游戏主逻辑
let board;
let currentBlock;
let gameInterval;
let score = 0;

export function init() {
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
    });
}

export function startGame() {
    score = 0;
    currentBlock = createBlock(board);
    renderBoard();
    
    // 添加键盘事件监听
    document.addEventListener("keydown", handleKeyDown);
    
    if (gameInterval) clearInterval(gameInterval);
    gameInterval = setInterval(() => {
        if (!currentBlock.down()) {
            board.placeBlock(currentBlock);
            const linesCleared = board.clearLines();
            score += linesCleared * 100;
            currentBlock = createBlock(board);
            
            // 检查游戏结束
            if (!board.isValidPosition(currentBlock.space)) {
                clearInterval(gameInterval);
                document.removeEventListener("keydown", handleKeyDown);
                console.log(`游戏结束！得分: ${score}`);
            }
        }
        renderBoard();
    }, 500);
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

export function blockMove(direction) {
    if (!board || !currentBlock) return;
    
    Excel.run(async (context) => {
        switch(direction) {
            case "rotate":
                currentBlock.rotate();
                break;
            case "left":
                currentBlock.left();
                break;
            case "right":
                currentBlock.right();
                break;
            case "down":
                currentBlock.down();
                break;
        }
        renderBoard();
        await context.sync();
    });
}

function renderBoard() {
    if (!board || !currentBlock) return;
    
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
    });
}
