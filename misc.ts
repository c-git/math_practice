const spreadsheet_active = SpreadsheetApp.getActive();

function getRandomNum(max?: number): number {
    if (isUndefined(max))
        max = getMAX_NUM();
    return getRandomInt(getMIN_NUM(), max);
}

function getRandomInt(min, max): number {
    return Math.floor(Math.random() * (max - min)) + min;
}

function getRandomArrayIndex(arr) {
    return getRandomInt(0, arr.length);
}

function getColNums(count: number): number[][] {
    let result = new Array(count);
    for (let i = 0; i < result.length; i++) {
        result[i] = [getRandomNum()];
    }
    return result;
}

function getEnabledArithOpers() {
    let result = getNamedRangeValues("EnabledArithOpers");

    // Remove Disabled functions
    for (let i = result.length - 1; i >= 0; i--) {
        if (!result[i][1]) {
            result.splice(i, 1);
        }
    }

    // Set generation functions
    for (let i = 0; i < result.length; i++) {
        if (result[i][0] === getOP_SYM_ADD()) {
            result[i][2] = genAdd;
        } else if (result[i][0] === getOP_SYM_SUB()) {
            result[i][2] = genSub;
        } else if (result[i][0] === getOP_SYM_MUL()) {
            result[i][2] = genMul;
        } else if (result[i][0] === getOP_SYM_DIV()) {
            result[i][2] = genDiv;
        } else
            throw "Unknown Operation" + result[i][0];
    }

    return result;
}

function getEnabledSkipBases() {
    let result = [];
    const bases = getSKIP_BASES();

    // Add Enabled Bases
    for (let i = 0; i < bases.length; i++) {
        if (bases[i][1]) {
            result.push(bases[i][0]);
        }
    }

    if (result.length === 0) {
        throw "No bases for Skip Counting Appear to be selected";
    }
    return result;
}

function genAdd(a, b) {
    return [a, getOP_SYM_ADD(), b];
}

function genSub(a, b) {
    return [a, getOP_SYM_SUB(), (a >= b) ? b : getRandomNum(a)];
}

function genMul(a, b) {
    return [a, getOP_SYM_MUL(), b];
}

function genDiv(a, b) {
    // Values limits
    // a = b to b * max
    // b = max
    // ans = max
    if (b === 0)
        b = 1;
    a = getRandomNum() * b;
    if (a === 0)
        a = b;
    return [a, getOP_SYM_DIV(), b];
}

function isUndefined(obj) {
    return typeof obj === "undefined";
}