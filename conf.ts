var getMIN_NUM = once(function () {
    return getNamedRangeValue("MinNum");
});
var getMAX_NUM = once(function () {
    return getNamedRangeValue("MaxNum") + 1;
});
var getQUESTION_COUNT = once(function () {
    return getNamedRangeValue("QuestionCount");
});
var getOP_SYM_ADD = once(function () {
    return getNamedRangeValue("OperAdd");
});
var getOP_SYM_SUB = once(function () {
    return getNamedRangeValue("OperMinus");
});
var getOP_SYM_MUL = once(function () {
    return getNamedRangeValue("OperMultiply");
});
var getOP_SYM_DIV = once(function () {
    return getNamedRangeValue("OperDivide");
});
var getSKIP_INCLUDE_FIRST = once(function () {
    return getNamedRangeValue("SkipCount_IncludeFirst");
});
var getSKIP_INCLUDE_RANDOM = once(function () {
    return getNamedRangeValue("SkipCount_IncludeRandom");
});
var getSKIP_RANDOM_PROB = once(function () {
    return getNamedRangeValue("SkipCount_Prob");
});
var getSKIP_LIMIT = once(function () {
    return getNamedRangeValue("SkipCount_Limit");
});
var getSKIP_BASES = once(function () {
    return getNamedRangeValues("SkipCount_Bases");
});

var COL_MIN = 2;
var COL_MAX = 4;

var ROW_NEW = 2;
var ROW_CLEAR = 3;
var ROW_MIN = ROW_NEW;
var ROW_MAX = ROW_CLEAR;

//////////////////// SUPPORTING FUNCTIONS ////////////////

function getNamedRangeValue(name) {
    return spreadsheet_active.getRange(name).getValue();
}

function getNamedRangeValues(name) {
    return spreadsheet_active.getRange(name).getValues();
}

function once(fn, context) {
    var result;

    return function () {
        if (fn) {
            result = fn.apply(context || this, arguments);
            fn = null;
        }

        return result;
    };
}
