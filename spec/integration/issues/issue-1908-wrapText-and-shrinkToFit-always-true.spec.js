const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 1908 - wrapText alignment properties always true', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(
      './spec/integration/data/test-issue-1908.xlsx'
    );

    // const [worksheet] = workbook.worksheets;

    workbook.eachSheet(sheet => {
      sheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, cellNumber) => {
          // eslint-disable-next-line no-console
          console.log({rowNumber, cellNumber});
          if (rowNumber === 1 && cellNumber === 1) {
            expect(cell.alignment.wrapText).to.equal(false);
          } else {
            expect(cell.alignment.wrapText).to.equal(true);
          }
        });
      });
    });

    // expect(worksheet.getCell('A1').alignment.wrapText).to.equal(false);
    // expect(worksheet.getCell('A2').alignment.wrapText).to.equal(true);
  });

  it('issue 1908 - wrapText alignment properties always true ..', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(
      './spec/integration/data/test-issue-1908.xlsx'
    );

    const [worksheet] = workbook.worksheets;

    expect(worksheet.getCell('A1').alignment.wrapText).to.equal(false);
    expect(worksheet.getCell('A2').alignment.wrapText).to.equal(true);
  });
});
