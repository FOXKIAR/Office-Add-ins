import { createBlock } from './blocks';

export function InitInterface() { // 初始化界面
    Excel.run(async (context) => {
        const workSheet = context.workbook.worksheets.getItem("Sheet1");
        const gameRange = workSheet.getRange("A1:L22");
        gameRange.format.fill.color = "black";
        gameRange.format.columnWidth = 15;
        await context.sync();
        const playRange = workSheet.getRange("B2:K21");
        playRange.format.fill.color = "white";
        await context.sync();
    });
}

export function playGame() {
    let block = createBlock();
    document.addEventListener("keydown", function (event) {
        if (event.code === "ArrowLeft") 
            blockMove(block, "left");
        if (event.code === "ArrowRight") 
            blockMove(block, "right");
    });

    setInterval(() => {
        blockMove(block, "down");
    }, 500)
}

export function blockMove(block, direction) {
    Excel.run(async (context) => {
        const workSheet = context.workbook.worksheets.getItem("Sheet1");
        const playRange = workSheet.getRange("B2:K21");
        block.space.forEach(item => {
            playRange.getCell(item.y, item.x).format.fill.color = "white";
        });
        if (direction === "left")
            block.left();
        if (direction === "right")
            block.right();
        if (direction === "down")
            block.down();

        block.space.forEach(item => {
            playRange.getCell(item.y, item.x).format.fill.color = block.color;
        });
        await context.sync();
    });
}