import type { Workbook } from '@mog-sdk/contracts/api';

export type EditScenarioId =
  | 'smoke-header-edit'
  | 'table-header-a1-only'
  | 'table-header-all-strings'
  | 'table-header-blank-cell'
  | 'table-header-duplicate-cell'
  | 'table-header-row-values'
  | 'table-header-formula-cell'
  | 'table-header-number-cell'
  | 'table-header-special-chars'
  | 'table-header-offset-row'
  | 'table-second-header-row'
  | 'table-data-row'
  | 'table-totals-row'
  | 'outside-table'
  | 'autofilter-header-row'
  | 'merged-anchor'
  | 'validation-cell'
  | 'hyperlink-cell'
  | 'defined-name-cell'
  | 'hidden-row-cell'
  | 'formula-cell'
  | 'shared-formula-child-cell'
  | 'legacy-array-input-cell'
  | 'dynamic-array-spill-cell'
  | 'formula-cache-cell'
  | 'calc-chain-formula-cell';

async function setHeaderRow(
  workbook: Workbook,
  startCell: 'A1' | 'B2' | 'F1' = 'A1',
) {
  const sheet = workbook.activeSheet;
  if (startCell === 'B2') {
    await sheet.setCell('B2', 'Mog E2E export smoke test');
    await sheet.setCell('C2', new Date().toISOString());
    await sheet.setCell('D2', '=LEN(B2)');
    return;
  }
  if (startCell === 'F1') {
    await sheet.setCell('F1', 'Mog E2E export smoke test');
    await sheet.setCell('G1', new Date().toISOString());
    await sheet.setCell('H1', '=LEN(F1)');
    return;
  }
  await sheet.setCell('A1', 'Mog E2E export smoke test');
  await sheet.setCell('B1', new Date().toISOString());
  await sheet.setCell('C1', '=LEN(A1)');
}

export async function applyEditScenario(workbook: Workbook, scenarioId: string) {
  await workbook.batch(`E2E ${scenarioId}`, async (batchedWorkbook) => {
    const sheet = batchedWorkbook.activeSheet;

    switch (scenarioId as EditScenarioId) {
      case 'smoke-header-edit':
      case 'table-header-row-values':
      case 'autofilter-header-row':
        await setHeaderRow(batchedWorkbook);
        return;
      case 'table-header-a1-only':
        await sheet.setCell('A1', 'Mog E2E renamed first header');
        return;
      case 'table-header-all-strings':
        await sheet.setCell('A1', 'Product');
        await sheet.setCell('B1', 'Segment');
        await sheet.setCell('C1', 'Quantity');
        await sheet.setCell('D1', 'Revenue');
        return;
      case 'table-header-blank-cell':
        await sheet.setCell('B1', '');
        return;
      case 'table-header-duplicate-cell':
        await sheet.setCell('B1', 'SKU');
        return;
      case 'table-header-formula-cell':
        await sheet.setCell('C1', '=LEN(A1)');
        return;
      case 'table-header-number-cell':
        await sheet.setCell('C1', 123);
        return;
      case 'table-header-special-chars':
        await sheet.setCell('A1', 'SKU[old]');
        await sheet.setCell('B1', 'Category#All');
        await sheet.setCell('C1', 'Units, Net');
        return;
      case 'table-header-offset-row':
        await setHeaderRow(batchedWorkbook, 'B2');
        return;
      case 'table-second-header-row':
        await setHeaderRow(batchedWorkbook, 'F1');
        return;
      case 'table-data-row':
        await sheet.setCell('A2', 'MOG-DATA-ROW');
        await sheet.setCell('B2', 'Edited data');
        await sheet.setCell('C2', 42);
        return;
      case 'table-totals-row':
        await sheet.setCell('A5', 'Edited total');
        await sheet.setCell('C5', 42);
        await sheet.setCell('D5', '=SUM(D2:D4)');
        return;
      case 'outside-table':
        await sheet.setCell('G7', 'Outside table edit');
        await sheet.setCell('H7', 123);
        return;
      case 'merged-anchor':
        await sheet.setCell('A1', 'Edited merged title');
        await sheet.setCell('A3', 123);
        return;
      case 'validation-cell':
        await sheet.setCell('B2', 'Rejected?');
        await sheet.setCell('C2', 999);
        return;
      case 'hyperlink-cell':
        await sheet.setCell('A2', 'Edited linked text');
        await sheet.setCell('B2', 'https://example.com/edited');
        return;
      case 'defined-name-cell':
        await sheet.setCell('B2', 777);
        await sheet.setCell('C2', '=NamedInput*2');
        return;
      case 'hidden-row-cell':
        await sheet.setCell('A3', 'Edited hidden-row-adjacent value');
        await sheet.setCell('B5', 555);
        return;
      case 'formula-cell':
        await sheet.setCell('C2', '=A2+B2');
        await sheet.setCell('D2', '=SUM(A2:C2)');
        return;
      case 'shared-formula-child-cell':
        await sheet.setCell('D3', '=B3*C3+1');
        await sheet.setCell('B4', 66);
        return;
      case 'legacy-array-input-cell':
        await sheet.setCell('A3', 20);
        await sheet.setCell('B5', 400);
        return;
      case 'dynamic-array-spill-cell':
        await sheet.setCell('C3', 'blocks spill');
        await sheet.setCell('A2', 10);
        return;
      case 'formula-cache-cell':
        await sheet.setCell('C2', '=A2-B2');
        await sheet.setCell('D2', '=TEXT(C2,"0")');
        return;
      case 'calc-chain-formula-cell':
        await sheet.setCell('C3', '=A3*B3');
        await sheet.setCell('D3', '=C3+100');
        return;
      default:
        throw new Error(`Unknown edit scenario: ${scenarioId}`);
    }
  });
}
