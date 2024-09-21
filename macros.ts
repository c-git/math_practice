/* global SpreadsheetApp, COL_MIN, COL_MAX, ROW_MIN, ROW_MAX, ROW_NEW, ROW_CLEAR */

/** @OnlyCurrentDoc */

var CMD_NEW = [NewEquations, CreateComparisons, NewSkipCounting];
var CMD_CLEAR = [ClearEquations, ClearComparison, ClearSkipCounting];

function ClearComparison() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comparison");
    spreadsheet.getRange('A:C').clear({ contentsOnly: true, skipFilteredRows: true });
}

function ClearEquations() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Equations");
    spreadsheet.getRange('A:C').clear({ contentsOnly: true, skipFilteredRows: true });
    spreadsheet.getRange('E:E').clear({ contentsOnly: true, skipFilteredRows: true });
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
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Skip Counting");
    spreadsheet.getRange('A1:A100').clear({ contentsOnly: true, skipFilteredRows: true });
}


function NewSkipCounting() {
    ClearSkipCounting();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Skip Counting");

    // Set new skip base
    var bases = getEnabledSkipBases();
    var base = bases[Math.floor(Math.random() * bases.length)];
    spreadsheet.getRange('B1').setValue(base);

    var given_values = new Array(getSKIP_LIMIT());

    // Set first element in skip counting (if enabled)
    given_values[0] = (getSKIP_INCLUDE_FIRST()) ? [base] : [null];

    // Set some other elements (if enabled)
    for (var i = 1; i < getSKIP_LIMIT(); i++) {
        given_values[i] =
            (getSKIP_INCLUDE_RANDOM() &&
                Math.random() < getSKIP_RANDOM_PROB())
                ? [(i + 1) * base] : [null];
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
        (e.range.getColumn() <= COL_MAX) &&
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
        } catch (err) {
            status_cell.setValue('ERROR: ' + err.message);
        }
    }
}