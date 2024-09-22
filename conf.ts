const getMIN_NUM = once(function () {
    return getNamedRangeValue("MinNum");
});
const getMAX_NUM = once(function () {
    return getNamedRangeValue("MaxNum") + 1;
});
const getQUESTION_COUNT = once(function () {
    return getNamedRangeValue("QuestionCount");
});
const getOP_SYM_ADD = once(function () {
    return getNamedRangeValue("OperAdd");
});
const getOP_SYM_SUB = once(function () {
    return getNamedRangeValue("OperMinus");
});
const getOP_SYM_MUL = once(function () {
    return getNamedRangeValue("OperMultiply");
});
const getOP_SYM_DIV = once(function () {
    return getNamedRangeValue("OperDivide");
});
const getSKIP_INCLUDE_FIRST = once(function () {
    return getNamedRangeValue("SkipCount_IncludeFirst");
});
const getSKIP_INCLUDE_RANDOM = once(function () {
    return getNamedRangeValue("SkipCount_IncludeRandom");
});
const getSKIP_RANDOM_PROB = once(function () {
    return getNamedRangeValue("SkipCount_Prob");
});
const getSKIP_LIMIT = once(function () {
    return getNamedRangeValue("SkipCount_Limit");
});
const getSKIP_BASES = once(function () {
    return getNamedRangeValues("SkipCount_Bases");
});
const getSKIP_IS_BASE_START = once(function () {
    return getNamedRangeValue("SkipCount_IsBaseStart");
});
const getSORT_COUNT = once(function () {
    return getNamedRangeValue("SortCount");
});
const getWORD_FORM_PROB = once(function () {
    return getNamedRangeValue("WordFormProb");
});
const getROUND_EXP_LOW = once(function () {
    return getNamedRangeValue("RoundExpLow");
});
const getROUND_EXP_HIGH = once(function () {
    return getNamedRangeValue("RoundExpHigh");
});

const COL_MIN = 2;

const ROW_NEW = 2;
const ROW_CLEAR = 3;
const ROW_MIN = ROW_NEW;
const ROW_MAX = ROW_CLEAR;

//////////////////// SUPPORTING FUNCTIONS ////////////////

function getNamedRangeValue(name: string) {
    return spreadsheet_active.getRange(name).getValue();
}

function getNamedRangeValues(name: string) {
    return spreadsheet_active.getRange(name).getValues();
}

function once(fn, context?) {
    let result;

    return function () {
        if (fn) {
            result = fn.apply(context || this, arguments);
            fn = null;
        }

        return result;
    };
}
