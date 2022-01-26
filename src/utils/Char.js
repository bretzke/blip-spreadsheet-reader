module.exports = class Char {
    static previousChar(char) {
        const c = char.toUpperCase();
        return `${
            c.substr(0, 1) === 'A' ? 'Z' : String.fromCharCode(c.charCodeAt(0) - 1)
        }${char.replace(/^./, '')}`;
    }

    static nextChar(char) {
        const c = char.toUpperCase();
        return `${
            c.substr(0, 1) === 'Z' ? 'A' : String.fromCharCode(c.charCodeAt(0) + 1)
        }${char.replace(/^./, '')}`;
    }
};
