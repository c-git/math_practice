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

function NewRounding() {
    ClearRounding();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rounding");
    const exponent_of_10 = getRandomInt(getROUND_EXP_LOW(), getROUND_EXP_HIGH());
    const round_to = 10 ** exponent_of_10;
    const nums_to_round = getColNums(getQUESTION_COUNT());
    const answers = [
        nums_to_round[0].map(x => Math.round(x / round_to) * round_to),
    ];
    spreadsheet.getRange('B1').setValue(round_to);
    spreadsheet.getRange('A3:A' + (3 + nums_to_round.length)).setValues(nums_to_round);
    spreadsheet.getRange('C3:C' + (3 + nums_to_round.length)).setValues(answers);
}

function NewWordForm() {
    ClearWordForm();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Word form");
    // TODO 1: Write func
}

function NewSort() {
    ClearSort();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sorting");
    // TODO 1: Write func
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
    const count = getQUESTION_COUNT();
    const l_col = getColNums(count);
    const r_col = getColNums(count);

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
    let nums = getColNums(getQUESTION_COUNT());
    spreadsheet.getRange('A1:A' + getQUESTION_COUNT()).setValues(nums);

    nums = getColNums(getQUESTION_COUNT());
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
    spreadsheet.getRange('A2:B100').clearContent();
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