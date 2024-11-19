/* global Excel console */

export async function insertText(text: string) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the current selected range
      const selection = context.workbook.getSelectedRange();
      selection.load(["address", "rowCount", "columnCount"]); // Load range details

      await context.sync();

      // Log the address of the selection
      console.log("Selected range:", selection.address);

      // Create a 2D array with the text repeated for each cell in the selection
      const values = Array(selection.rowCount)
        .fill(null)
        .map(() => Array(selection.columnCount).fill(text));

      // Set values in the selected range
      selection.values = values;

      // Auto-fit columns
      selection.format.autofitColumns();

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

/* global Excel console */

export async function sumSelectedRange() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the current selected range
      const selection = context.workbook.getSelectedRange();
      selection.load(["values", "rowIndex", "rowCount", "columnIndex"]);

      await context.sync();

      // console.log("Selected range:", selection.address);

      // Calculate the sum of the selected values
      let sum = 0;
      selection.values.forEach((row: any[]) => {
        row.forEach((cell) => {
          if (!isNaN(cell)) {
            sum += parseFloat(cell); // Accumulate only numeric values
          }
        });
      });

      console.log("Sum of selected values:", sum);

      // Determine the next cell below the selection to display the result
      const startRow = selection.rowIndex; // Starting row index
      const endRow = startRow + selection.rowCount; // Next row after selection
      const startColumn = selection.columnIndex; // Starting column index
      const resultCell = sheet.getCell(endRow, startColumn); // First column below selection

      // Display the sum in the result cell
      resultCell.values = [[sum]];
      resultCell.format.autofitColumns();

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

