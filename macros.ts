/** @OnlyCurrentDoc */

// Version: 2.0

const CMD_NEW = [
    NewEquations,
    CreateComparisons,
    NewSkipCounting,
    NewRounding,
    NewWordForm,
    NewSort
];
const CMD_CLEAR = [
    ClearEquations,
    ClearComparison,
    ClearSkipCounting,
    ClearRounding,
    ClearWordForm,
    ClearSort
];

/**
 * Converts numbers into words
 * 
 * Taken from https://stackoverflow.com/questions/72159705/numbers-to-words-using-html-and-javascript
 * @param n number to be converted
 * @returns String representation of the number given
 */
function number_to_words(n: number): string {
    const num = "zero,one,two,three,four,five,six,seven,eight,nine,ten,eleven,twelve,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,nineteen".split(",");
    const tens = "twenty,thirty,forty,fifty,sixty,seventy,eighty,ninety".split(",");
    if (n < 20) return num[n];
    const digit = n % 10;
    if (n < 100) return tens[Math.floor((n / 10)) - 2] + (digit ? "-" + num[digit] : "");
    if (n < 1000) return num[Math.floor((n / 100))] + " hundred" + (n % 100 == 0 ? "" : " and " + number_to_words(n % 100));
    if (n < 1000000) return number_to_words(Math.floor((n / 1000))) + " thousand" + (n % 1000 == 0 ? "" : ((n % 1000 < 100 ? " and " : " ") + number_to_words(n % 1000)));
    return number_to_words(Math.floor((n / 1000000))) + " million" + (n % 1000000 == 0 ? "" : ((n % 1000000 < 100 ? " and " : " ") + number_to_words(n % 1000000)));
}

function NewRounding() {
    ClearRounding();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rounding");
    const exponent_of_10 = getRandomInt(getROUND_EXP_LOW(), getROUND_EXP_HIGH());
    const round_to = 10 ** exponent_of_10;
    const nums_to_round = getColNums();
    spreadsheet.getRange('B1').setValue(round_to);
    spreadsheet.getRange('A3:A' + (3 - 1 + nums_to_round.length)).setValues(nums_to_round);
}

function NewWordForm() {
    ClearWordForm();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Word form");
    let nums = getColNums();
    const output = new Array(getQUESTION_COUNT());
    for (let i = 0; i < output.length; i++) {
        const ground_truth = nums[i][0];
        output[i] = (Math.random() < getWORD_FORM_PROB() ? [ground_truth, "", ground_truth] : ["", number_to_words(ground_truth), ground_truth]);
    }
    spreadsheet.getRange('A1:C' + getQUESTION_COUNT()).setValues(output);
}

function NewSort() {
    ClearSort();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sorting");
    const count = getSORT_COUNT();
    const nums = getColNums(count);
    spreadsheet.getRange('A1:A' + count).setValues(nums);
}

function ClearComparison() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comparison");
    spreadsheet.getRange('A:C').clearContent();
}

function ClearEquations() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Equations");
    spreadsheet.getRange('A:C').clearContent();
    spreadsheet.getRange('E:E').clearContent();
}

function NewEquations() {
    ClearEquations();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Equations");

    // Get default values for left and right column
    const l_col = getColNums();
    const r_col = getColNums();

    const ops = getEnabledArithOpers();
    const eqs = new Array(getQUESTION_COUNT());
    for (let i = 0; i < eqs.length; i++) {
        eqs[i] = ops[getRandomArrayIndex(ops)][2](l_col[i][0], r_col[i][0]);
    }

    spreadsheet.getRange('A1:C' + getQUESTION_COUNT()).setValues(eqs);
}

function CreateComparisons() {
    ClearComparison();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comparison");
    let nums = getColNums();
    spreadsheet.getRange('A1:A' + getQUESTION_COUNT()).setValues(nums);

    nums = getColNums();
    spreadsheet.getRange('C1:C' + getQUESTION_COUNT()).setValues(nums);
}

function ClearSkipCounting() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Skip Counting");
    spreadsheet.getRange('A1:A100').clearContent();
}

function ClearSort() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sorting");
    spreadsheet.getRange('A1:A30').clearContent();
}

function ClearRounding() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rounding");
    spreadsheet.getRange('B3:B100').clearContent();
}

function ClearWordForm() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Word form");
    spreadsheet.getRange('A2:C100').clearContent();
}


function NewSkipCounting() {
    ClearSkipCounting();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Skip Counting");

    // Set new skip base
    const bases = getEnabledSkipBases();
    const base_increment = bases[Math.floor(Math.random() * bases.length)];
    const start_val = (getSKIP_IS_BASE_START()) ? base_increment : getRandomNum();
    spreadsheet.getRange('B1:C1').setValues([[start_val, base_increment]]);

    const given_values = new Array(getSKIP_LIMIT());

    // Set first element in skip counting (if enabled)
    given_values[0] = (getSKIP_INCLUDE_FIRST()) ? [start_val] : [null];

    // Set some other elements (if enabled)
    for (let i = 1; i < getSKIP_LIMIT(); i++) {
        given_values[i] =
            (getSKIP_INCLUDE_RANDOM() &&
                Math.random() < getSKIP_RANDOM_PROB())
                ? [start_val + i * base_increment] : [null];
    }

    // Write to sheet
    spreadsheet.getRange('A1:A' + getSKIP_LIMIT()).setValues(given_values);
}

function InstalledEdit(e) {
    if (
        (e.source.getActiveSheet().getName() === "Controls") &&
        (e.range.getWidth() === 1) &&
        (e.range.getHeight() === 1) &&
        (e.range.getColumn() >= COL_MIN) &&
        (e.range.getRow() >= ROW_MIN) &&
        (e.range.getRow() <= ROW_MAX)
    ) {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls");
        const status_cell = spreadsheet.getRange('A6');
        try {
            if (status_cell.getValue() === 'Done') {
                status_cell.setValue('Running...');
            } else {
                Logger.log('Aborted status not done');
                return; // Abort previous status was not done
            }
            if (e.range.getRow() === ROW_NEW) {
                e.range.setValue('NEW');
                CMD_NEW[e.range.getColumn() - COL_MIN]();
            } else if (e.range.getRow() === ROW_CLEAR) {
                e.range.setValue('CLEAR');
                CMD_CLEAR[e.range.getColumn() - COL_MIN]();
            }
            status_cell.setValue('Done');
            Logger.log('Done');
        } catch (err) {
            const err_msg = `ERROR: ${err.message}`;
            status_cell.setValue(err_msg);
            Logger.log(err_msg);
        }
    }
}