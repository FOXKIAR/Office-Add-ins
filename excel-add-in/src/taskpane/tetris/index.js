import blocks from './blocks';

export async function InitInterface() { // 初始化界面
    Excel.run(async (context) => {
        const workSheet = context.workbook.worksheets.getItem("Sheet1");
        const gameRange = workSheet.getRange("A1:L22");
        gameRange.format.fill.color = "black";
        gameRange.format.columnWidth = 15;
        await context.sync();
        const playRange = workSheet.getRange("B2:K21");
        playRange.format.fill.color = "white";
        createBlock(playRange);
        await context.sync();
    });
}

export function createBlock(range) { // 初始化方块信息
    const block = blocks[Math.floor(Math.random() * 7)];
    for (let i = 0; i < 4; i ++) {
        range.getCell(block.space[i].y, block.space[i].x).format.fill.color = block.color;
    }
}

