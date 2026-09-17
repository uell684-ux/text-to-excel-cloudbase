const ExcelJS = require("exceljs");

/**
 * Convert structured Dify data into an Excel workbook buffer.
 *
 * Expected data:
 * {
 *   title: "...",
 *   columns: [...],
 *   rows: [...],
 *   summary: [...]
 * }
 */
async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Text to Excel";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("数据");

  const title = data.title || "数据整理结果";
  const columns = Array.isArray(data.columns) ? data.columns : [];
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const summary = Array.isArray(data.summary) ? data.summary : [];

  if (columns.length === 0) {
    throw new Error("No columns found");
  }

  // Title
  worksheet.mergeCells(1, 1, 1, columns.length);

  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = {
    bold: true,
    size: 16
  };
  titleCell.alignment = {
    horizontal: "center",
    vertical: "middle"
  };

  worksheet.getRow(1).height = 28;

  // Header
  const headerRow = worksheet.getRow(3);

  columns.forEach((column, index) => {
    headerRow.getCell(index + 1).value = column;
  });

  headerRow.font = {
    bold: true
  };

  headerRow.alignment = {
    horizontal: "center",
    vertical: "middle"
  };

  // Data rows
  rows.forEach((row) => {
    const values = [];

    for (let i = 0; i < columns.length; i++) {
      values.push(row[i] !== undefined ? row[i] : "");
    }

    worksheet.addRow(values);
  });

  // Summary
  if (summary.length > 0) {
    worksheet.addRow([]);

    summary.forEach((item) => {
      const summaryRow = new Array(columns.length).fill("");

      summaryRow[0] = item.label || "总计";

      const values = item.values || {};

      columns.forEach((column, index) => {
        if (Object.prototype.hasOwnProperty.call(values, column)) {
          summaryRow[index] = values[column];
        }
      });

      const row = worksheet.addRow(summaryRow);

      row.font = {
        bold: true
      };
    });
  }

  // Borders + alignment
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= 3) {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };

        cell.alignment = {
          vertical: "middle",
          wrapText: true
        };
      });
    }
  });

  // Dynamic column width
  columns.forEach((column, index) => {
    let maxLength = String(column).length;

    rows.forEach((row) => {
      const value = row[index];

      if (value !== undefined && value !== null) {
        maxLength = Math.max(
          maxLength,
          String(value).length
        );
      }
    });

    worksheet.getColumn(index + 1).width =
      Math.min(Math.max(maxLength + 4, 12), 35);
  });

  worksheet.views = [
    {
      state: "frozen",
      ySplit: 3
    }
  ];

  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  generateExcel
};
