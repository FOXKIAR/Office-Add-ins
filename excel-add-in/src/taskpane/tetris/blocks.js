class Block {
    space = [{x: 0, y: 0}, {x: 0, y: 0}, {x: 0, y: 0}, {x: 0, y: 0}];
    color = "";

    down() {
        this.space.forEach((item) => item.y += 1);
    }

    left() {
        this.space.forEach((item) => item.x -= 1);
    }

    right() {
        this.space.forEach((item) => item.x += 1);
    }
}


// ■ ■ ■
//   ■
class T extends Block {
    space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 6, y: 0}, {x: 5, y: 1}]; 
    color = "yellow";
}

// ■
// ■
// ■ ■
class L extends Block {
    space = [{x: 4, y: 0}, {x: 4, y: 1}, {x: 4, y: 2}, {x: 5, y: 2}]; 
    color = "red";
}

//   ■
//   ■
// ■ ■
class J extends Block {
    space = [{x: 5, y: 0}, {x: 5, y: 1}, {x: 4, y: 2}, {x: 5, y: 2}]; 
    color = "red";
}

// ■ ■
//   ■ ■
class Z extends Block {
    space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 5, y: 1}, {x: 6, y: 1}]; 
    color = "green";
}

//   ■ ■
// ■ ■
class S extends Block {
    space = [{x: 5, y: 0}, {x: 6, y: 0}, {x: 4, y: 1}, {x: 5, y: 1}]; 
    color = "green";
}

// ■ ■
// ■ ■
class O extends Block {
    space = [{x: 4, y: 0}, {x: 5, y: 0}, {x: 4, y: 1}, {x: 5, y: 1}]; 
    color = "blue";
}

// ■
// ■
// ■ 
// ■
class I extends Block {
    space = [{x: 5, y: 0}, {x: 5, y: 1}, {x: 5, y: 2}, {x: 5, y: 3}]; 
    color = "blue";
}

export function createBlock() {
    switch(Math.floor(Math.random() * 7)) {
        case 0: 
            return new T();
        case 1:
            return new L();
        case 2: 
            return new J();
        case 3:
            return new Z();
        case 4: 
            return new S();
        case 5:
            return new O();
        case 6: 
            return new I();
    }
}