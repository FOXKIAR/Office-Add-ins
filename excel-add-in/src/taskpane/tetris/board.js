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
        for (let y = this.height - 1; y >= 0; y--) {
            if (this.grid[y].every(cell => cell !== null)) {
                this.grid.splice(y, 1);
                this.grid.unshift(Array(this.width).fill(null));
                linesCleared++;
                y++;
            }
        }
        return linesCleared;
    }
}