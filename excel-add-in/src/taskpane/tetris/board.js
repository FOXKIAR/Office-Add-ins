export class Board {
    constructor(width = 10, height = 20) {
        this.width = width;
        this.height = height;
        this.grid = Array(height).fill().map(() => Array(width).fill(null));
    }

    isValidPosition(positions) {
        return positions.every(pos => 
            pos.x >= 0 && 
            pos.x < this.width && 
            pos.y < this.height && 
            (pos.y < 0 || !this.grid[pos.y][pos.x])
        );
    }

    placeBlock(block) {
        block.space.forEach(pos => {
            if (pos.y >= 0) {
                this.grid[pos.y][pos.x] = block.color;
            }
        });
    }

    clearLines() {
        let linesCleared = 0;
        // 从底部向上检查，这样删除行时不会影响未检查的行
        for (let y = this.height - 1; y >= 0; y--) {
            // 检查当前行是否已满
            if (this.grid[y].every(cell => cell !== null)) {
                // 删除满行并在顶部添加空行
                this.grid.splice(y, 1);
                this.grid.unshift(Array(this.width).fill(null));
                linesCleared++;
                // y++ 确保重新检查同一位置（因为删除后索引会变化）
                // 这样可以清除连续的多行
                y++;
            }
        }
        return linesCleared;
    }
}