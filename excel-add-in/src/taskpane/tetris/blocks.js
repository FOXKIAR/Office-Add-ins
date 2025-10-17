class Block {
    constructor(board) {
        this.board = board;
        this.rotation = 0;
    }

    move(dx, dy) {
        const newSpace = this.space.map(pos => ({x: pos.x + dx, y: pos.y + dy}));
        if (this.board.isValidPosition(newSpace)) {
            this.space = newSpace;
            return true;
        }
        return false;
    }

    down() {
        return this.move(0, 1);
    }

    left() {
        return this.move(-1, 0);
    }

    right() {
        return this.move(1, 0);
    }

    rotate() {
        const pivot = this.getPivot();
        const newSpace = this.space.map(pos => {
            const relX = pos.x - pivot.x;
            const relY = pos.y - pivot.y;
            return {
                x: pivot.x - relY,
                y: pivot.y + relX
            };
        });
        
        if (this.board.isValidPosition(newSpace)) {
            this.space = newSpace;
            this.rotation = (this.rotation + 1) % 4;
            return true;
        }
        return false;
    }

    getPivot() {
        return this.space[1];
    }
}

class T extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 6, y: 0}, {x: 5, y: 1}];
        this.color = "yellow";
    }
}

class L extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 4, y: 0}, {x: 4, y: 1}, {x: 4, y: 2}, {x: 5, y: 2}];
        this.color = "red";
    }
}

class J extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 5, y: 0}, {x: 5, y: 1}, {x: 4, y: 2}, {x: 5, y: 2}];
        this.color = "blue";
    }
}

class Z extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 5, y: 1}, {x: 6, y: 1}];
        this.color = "green";
    }
}

class S extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 5, y: 0}, {x: 6, y: 0}, {x: 4, y: 1}, {x: 5, y: 1}];
        this.color = "purple";
    }
}

class O extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 4, y: 1}, {x: 5, y: 1}];
        this.color = "orange";
    }
    
    rotate() {
        return false; // O方块不需要旋转
    }
}

class I extends Block {
    constructor(board) {
        super(board);
        this.space = [{x: 5, y: 0}, {x: 5, y: 1}, {x: 5, y: 2}, {x: 5, y: 3}];
        this.color = "cyan";
    }
}

export function createBlock(board) {
    const blocks = [T, L, J, Z, S, O, I];
    const BlockClass = blocks[Math.floor(Math.random() * 7)];
    return new BlockClass(board);
}