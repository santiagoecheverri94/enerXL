const FIRST_COLUMN_IN_RANGE = 0;
const GOVT_ENGINE_SHEET_NAME = 'H2ALite';
const GOVT_ENGINE_INPUT_CELL = 'B4';
const GOVT_ENGINE_OUTPUT_CELL = 'C3';

export async function fillOutputColumnToTheRightOfInputRange(context: Excel.RequestContext) {
  const range = context.workbook.getSelectedRange();
  range.load('rowCount');
  await context.sync();

  // Iterate over each cell in the first column of the range
  for (let row = 0; row < range.rowCount; row++) {
    const realTimeLmpInputCell = range.getCell(row, FIRST_COLUMN_IN_RANGE);



    const realLevelizedCostOutputCell = getCellToTheRightOfRealTimeLmpInput(realTimeLmpInputCell);
    
    realLevelizedCostOutputCell.values = [[0]];
    await context.sync();
  }
}

function getCellToTheRightOfRealTimeLmpInput(inputCell: Excel.Range): Excel.Range {
  return inputCell.getOffsetRange(0, 1);
}
