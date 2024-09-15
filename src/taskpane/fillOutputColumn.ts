const FIRST_COLUMN_IN_RANGE = 0;
const GOVT_ENGINE_SHEET_NAME = 'H2ALite';
const GOVT_ENGINE_INPUT_CELL = 'G27';
const GOVT_ENGINE_OUTPUT_CELL = 'M1';

export async function fillOutputColumnToTheRightOfInputRange(context: Excel.RequestContext) {
  const range = context.workbook.getSelectedRange();
  range.load(['rowCount', 'values']);
  await context.sync();

  const realTimeLmpInputValues = range.values;

  // Iterate over each cell in the first column of the range
  for (let row = 0; row < range.rowCount; row++) {
    const realTimeLmpInputValue = realTimeLmpInputValues[row][FIRST_COLUMN_IN_RANGE];
    if (typeof realTimeLmpInputValue !== 'number') {
      continue;
    }

    const realLevelizedCostOutputValue = await getGovtEngineOutputValue(context, realTimeLmpInputValue);

    const realLevelizedCostOutputCell = getCellToTheRightOfRealTimeLmpInput(getRealTimeLmpInputCell(range, row));
    realLevelizedCostOutputCell.values = [[realLevelizedCostOutputValue]];
    await context.sync();
  }
}

function getRealTimeLmpInputCell(range: Excel.Range, row: number): Excel.Range {
  return range.getCell(row, FIRST_COLUMN_IN_RANGE);
}

async function getGovtEngineOutputValue(context: Excel.RequestContext, realTimeLmpInputValue: number): Promise<number> {
  const govtEngineSheet = context.workbook.worksheets.getItem(GOVT_ENGINE_SHEET_NAME);
  
  const govtEngineInputCell = govtEngineSheet.getRange(GOVT_ENGINE_INPUT_CELL);
  govtEngineInputCell.values = [[realTimeLmpInputValue]];
  
  let govtEngineOutputCell = govtEngineSheet.getRange(GOVT_ENGINE_OUTPUT_CELL);
  govtEngineOutputCell.calculate();

  do  {
    context.application.load('calculationState');
    await context.sync();
  } while(context.application.calculationState !== Excel.CalculationState.done);
  
  govtEngineOutputCell = govtEngineSheet.getRange(GOVT_ENGINE_OUTPUT_CELL);
  govtEngineOutputCell.load('values');
  await context.sync();

  return govtEngineOutputCell.values[0][0];
}

function getCellToTheRightOfRealTimeLmpInput(inputCell: Excel.Range): Excel.Range {
  return inputCell.getOffsetRange(0, 1);
}
