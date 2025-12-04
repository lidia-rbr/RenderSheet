/**
 * Mini "React-like" renderer for Google Sheets using batchUpdate.
 *
 * Prerequisites:
 * - Enable Advanced Sheets service: Services > Google Sheets API (Sheets).
 */


/**
 * Converts a 1-based column index to an A1 column letter (1 -> "A").
 *
 * @param {number} index 1-based column index.
 * @returns {string} Column letter.
 */
function columnIndexToLetter(index) {
  let temp = index;
  let letter = "";
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

/**
 * Central context that accumulates value updates and batchUpdate requests.
 */
class SheetContext {
  /**
   * @param {string} spreadsheetId Target spreadsheet ID.
   */
  constructor(spreadsheetId) {
    this.spreadsheetId = spreadsheetId;
    this.valueRanges = [];
    this.requests = [];
  }

  /**
   * Queues a values batch update entry for the given A1 range.
   *
   * @param {string} rangeA1 A1 notation, e.g. "Sheet1!A1:C3".
   * @param {unknown[][]} values Two-dimensional array of values.
   * @returns {void}
   */
  writeRange(rangeA1, values) {
    this.valueRanges.push({
      range: rangeA1,
      values,
    });
  }

  /**
   * Queues a background color change for a range.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range within the sheet, e.g. "A1:C3".
   * @param {string} color Hex color, e.g. "#0f172a".
   * @returns {void}
   */
  setBackground(sheetName, a1Range, color) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      repeatCell: {
        range: gridRange,
        cell: {
          userEnteredFormat: {
            backgroundColor: this.hexToRgbColor(color),
          },
        },
        fields: "userEnteredFormat.backgroundColor",
      },
    });
  }

  /**
   * Queues a number format update for a range.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {string} numberFormatPattern Number format pattern, e.g. "#,##0.00".
   * @returns {void}
   */
  setNumberFormat(sheetName, a1Range, numberFormatPattern) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      repeatCell: {
        range: gridRange,
        cell: {
          userEnteredFormat: {
            numberFormat: {
              type: "NUMBER",
              pattern: numberFormatPattern,
            },
          },
        },
        fields: "userEnteredFormat.numberFormat",
      },
    });
  }

  /**
   * Queues a text formatting update (bold, italic, color, size).
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {{
   *   bold?: boolean,
   *   italic?: boolean,
   *   underline?: boolean,
   *   fontSize?: number,
   *   color?: string
   * }} options Text format options.
   * @returns {void}
   */
  setTextFormat(sheetName, a1Range, options) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    /** @type {GoogleAppsScript.Sheets.Schema.TextFormat} */
    const textFormat = {};

    if (options.bold !== undefined) textFormat.bold = options.bold;
    if (options.italic !== undefined) textFormat.italic = options.italic;
    if (options.underline !== undefined) textFormat.underline = options.underline;
    if (options.fontSize !== undefined) textFormat.fontSize = options.fontSize;
    if (options.color) {
      textFormat.foregroundColor = this.hexToRgbColor(options.color);
    }

    this.requests.push({
      repeatCell: {
        range: gridRange,
        cell: {
          userEnteredFormat: {
            textFormat,
          },
        },
        fields: "userEnteredFormat.textFormat",
      },
    });
  }

  /**
   * Queues horizontal and/or vertical alignment update.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {{
   *   horizontal?: "LEFT" | "CENTER" | "RIGHT",
   *   vertical?: "TOP" | "MIDDLE" | "BOTTOM"
   * }} options Alignment options.
   * @returns {void}
   */
  setAlignment(sheetName, a1Range, options) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    const format = {};
    const fields = [];

    if (options.horizontal) {
      format.horizontalAlignment = options.horizontal;
      fields.push("userEnteredFormat.horizontalAlignment");
    }
    if (options.vertical) {
      format.verticalAlignment = options.vertical;
      fields.push("userEnteredFormat.verticalAlignment");
    }

    if (fields.length === 0) return;

    this.requests.push({
      repeatCell: {
        range: gridRange,
        cell: {
          userEnteredFormat: format,
        },
        fields: fields.join(","),
      },
    });
  }

  /**
   * Queues a wrap strategy update for a range.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {"WRAP" | "OVERFLOW_CELL" | "CLIP"} wrapStrategy Wrap strategy.
   * @returns {void}
   */
  setWrapStrategy(sheetName, a1Range, wrapStrategy) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      repeatCell: {
        range: gridRange,
        cell: {
          userEnteredFormat: {
            wrapStrategy,
          },
        },
        fields: "userEnteredFormat.wrapStrategy",
      },
    });
  }

  /**
   * Queues uniform borders on a range.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {{
   *   style?: "SOLID" | "DASHED" | "DOTTED" | "SOLID_THICK" | "DOUBLE",
   *   width?: number,
   *   color?: string
   * }} options Border options.
   * @returns {void}
   */
  setBorders(sheetName, a1Range, options) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    const style = options.style || "SOLID";
    const width = options.width || 1;
    const color = options.color ? this.hexToRgbColor(options.color) : undefined;

    const border = {
      style,
      width,
      color,
    };

    this.requests.push({
      updateBorders: {
        range: gridRange,
        top: border,
        bottom: border,
        left: border,
        right: border,
        innerHorizontal: border,
        innerVertical: border,
      },
    });
  }

  /**
   * Queues a mergeCells request.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {"MERGE_ALL" | "MERGE_ROWS" | "MERGE_COLUMNS"} mergeType Merge type.
   * @returns {void}
   */
  mergeRange(sheetName, a1Range, mergeType) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      mergeCells: {
        range: gridRange,
        mergeType: mergeType || "MERGE_ALL",
      },
    });
  }

  /**
   * Queues an unmergeCells request.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @returns {void}
   */
  unmergeRange(sheetName, a1Range) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      unmergeCells: {
        range: gridRange,
      },
    });
  }

  /**
   * Queues an auto-resize for all columns intersecting the range.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @returns {void}
   */
  autoResizeColumns(sheetName, a1Range) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange) return;

    this.requests.push({
      autoResizeDimensions: {
        dimensions: {
          sheetId: gridRange.sheetId,
          dimension: "COLUMNS",
          startIndex: gridRange.startColumnIndex,
          endIndex: gridRange.endColumnIndex,
        },
      },
    });
  }

  /**
   * Queues a dropdown data validation (ONE_OF_LIST).
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @param {string[]} values Allowed values.
   * @returns {void}
   */
  setDataValidationList(sheetName, a1Range, values) {
    const gridRange = this.convertA1ToGridRange(sheetName, a1Range);
    if (!gridRange || values.length === 0) return;

    this.requests.push({
      setDataValidation: {
        range: gridRange,
        rule: {
          condition: {
            type: "ONE_OF_LIST",
            values: values.map(value => ({ userEnteredValue: value })),
          },
          showCustomUi: true,
        },
      },
    });
  }

  /**
   * Queues a raw Sheets batchUpdate request.
   *
   * @param {SheetsRequest} request A Sheets API request.
   * @returns {void}
   */
  addRequest(request) {
    this.requests.push(request);
  }

  /**
   * Flushes all queued value updates and batchUpdate requests.
   *
   * @returns {void}
   */
  commit() {
    // 1) Run formatting / structural changes first
    if (this.requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate(
        {
          requests: this.requests,
        },
        this.spreadsheetId
      );
    }

    // 2) Then write all values so nothing can overwrite them afterwards
    if (this.valueRanges.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate(
        {
          valueInputOption: "USER_ENTERED",
          data: this.valueRanges,
        },
        this.spreadsheetId
      );
    }
  }

  /**
   * Converts a hex color string to a Sheets Color object (0–1 range).
   *
   * @param {string} hex Hex color string, e.g. "#ffffff".
   * @returns {GoogleAppsScript.Sheets.Schema.Color} Sheets Color object.
   */
  hexToRgbColor(hex) {
    let clean = hex.replace("#", "");
    if (clean.length === 3) {
      clean = clean
        .split("")
        .map(ch => ch + ch)
        .join("");
    }

    const r = parseInt(clean.substring(0, 2), 16) / 255;
    const g = parseInt(clean.substring(2, 4), 16) / 255;
    const b = parseInt(clean.substring(4, 6), 16) / 255;

    return {
      red: r,
      green: g,
      blue: b,
    };
  }

  /**
   * Converts an A1 range within a sheet into a GridRange.
   *
   * @param {string} sheetName Target sheet name.
   * @param {string} a1Range A1 range.
   * @returns {GoogleAppsScript.Sheets.Schema.GridRange | null} GridRange or null.
   */
  convertA1ToGridRange(sheetName, a1Range) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }

    const fullRange = sheet.getRange(a1Range);
    const sheetId = sheet.getSheetId();

    const startRow = fullRange.getRow() - 1;
    const endRow = startRow + fullRange.getNumRows();
    const startCol = fullRange.getColumn() - 1;
    const endCol = startCol + fullRange.getNumColumns();

    return {
      sheetId,
      startRowIndex: startRow,
      endRowIndex: endRow,
      startColumnIndex: startCol,
      endColumnIndex: endCol,
    };
  }
}

/**
 * Base class for a Sheet "component".
 */
class SheetComponent {
  /**
   * @param {SheetContext} context Shared SheetContext instance.
   * @param {Object} props Props passed to the component.
   */
  constructor(context, props) {
    this.context = context;
    this.props = props || {};
  }

  /**
   * Called when the component is rendered. Should be overridden by subclasses.
   *
   * @returns {void}
   */
  render() {

  }

  /**
   * Helper that instantiates and renders a child component.
   *
   * @param {typeof SheetComponent} ComponentClass Component class.
   * @param {Object<string, unknown>} props Props for the child.
   * @returns {SheetComponent} Child instance.
   */
  renderChild(ComponentClass, props) {
    const child = new ComponentClass(this.context, props);
    child.render();
    return child;
  }
}

/**
 * Renders a root component tree into a spreadsheet.
 *
 * @param {string} spreadsheetId Target spreadsheet ID.
 * @param {typeof SheetComponent} RootComponent Root component class.
 * @param {Object<string, unknown>} props Props for the root component.
 * @returns {void}
 */
function renderSheet(spreadsheetId, RootComponent, props) {
  const ctx = new SheetContext(spreadsheetId);
  const root = new RootComponent(ctx, props || {});
  root.render();
  ctx.commit();
}

/**
 * Component that writes a big title and optional subtitle at the top.
 */
class TitleComponent extends SheetComponent {

  render() {
    const sheetName = (this.props.sheetName);
    const title = (this.props.title);
    const subtitle = (this.props.subtitle);

    this.context.writeRange(`${sheetName}!A1`, [[title]]);
    this.context.mergeRange(sheetName, "A1:D1", "MERGE_ALL");
    this.context.setBackground(sheetName, "A1:D1", "#0f172a");
    this.context.setTextFormat(sheetName, "A1:D1", {
      bold: true,
      fontSize: 16,
      color: "#ffffff",
    });
    this.context.setAlignment(sheetName, "A1:D1", { horizontal: "CENTER", vertical: "MIDDLE" });

    if (subtitle) {
      this.context.writeRange(`${sheetName}!A2`, [[subtitle]]);
      this.context.mergeRange(sheetName, "A2:D2", "MERGE_ALL");
      this.context.setTextFormat(sheetName, "A2:D2", {
        italic: true,
        color: "#475569",
      });
      this.context.setAlignment(sheetName, "A2:D2", { horizontal: "CENTER", vertical: "MIDDLE" });
    }
  }
}

/**
 * Component that writes a header row with styling.
 */
class HeaderRowComponent extends SheetComponent {
  /**
   * @returns {void}
   */
  render() {
    const sheetName = (this.props.sheetName);
    const headers = (this.props.headers);
    const rowIndex = (this.props.rowIndex);

    const lastColLetter = columnIndexToLetter(headers.length);
    const rangeA1 = `${sheetName}!A${rowIndex}:${lastColLetter}${rowIndex}`;

    this.context.writeRange(rangeA1, [headers]);
    this.context.setBackground(sheetName, `A${rowIndex}:${lastColLetter}${rowIndex}`, "#1e293b");
    this.context.setTextFormat(sheetName, `A${rowIndex}:${lastColLetter}${rowIndex}`, {
      bold: true,
      color: "#ffffff",
    });
    this.context.setAlignment(sheetName, `A${rowIndex}:${lastColLetter}${rowIndex}`, {
      horizontal: "CENTER",
      vertical: "MIDDLE",
    });
  }
}

/**
 * Component that writes a 2D data table starting at a given row.
 */
class DataTableComponent extends SheetComponent {
  /**
   * @returns {void}
   */
  render() {
    const sheetName = (this.props.sheetName);
    const rows = (this.props.rows);
    const startRow = (this.props.startRow);

    if (!rows || rows.length === 0) return;

    const numCols = rows[0].length;
    const lastColLetter = columnIndexToLetter(numCols);
    const endRow = startRow + rows.length - 1;
    const rangeA1 = `${sheetName}!A${startRow}:${lastColLetter}${endRow}`;

    this.context.writeRange(rangeA1, rows);
    this.context.setBorders(sheetName, `A${startRow}:${lastColLetter}${endRow}`, {
      style: "SOLID",
      width: 1,
      color: "#cbd5e1",
    });
    this.context.setAlignment(sheetName, `A${startRow}:${lastColLetter}${endRow}`, {
      vertical: "MIDDLE",
    });
    this.context.setWrapStrategy(sheetName, `A${startRow}:${lastColLetter}${endRow}`, "WRAP");
  }
}

class SummaryRowComponent extends SheetComponent {
  /**
   * @returns {void}
   */
  render() {
    const sheetName = (this.props.sheetName);
    const label = (this.props.label);
    const rowIndex = (this.props.rowIndex);
    const numCols = (this.props.numCols);
    const amountColumnIndex = (this.props.amountColumnIndex);

    const lastColLetter = columnIndexToLetter(numCols);
    const labelCell = `A${rowIndex}`;
    const amountCell = `${columnIndexToLetter(numCols)}${rowIndex}`;
    const labelRangeSingle = `${sheetName}!${labelCell}`;
    const fullRowRange = `A${rowIndex}:${lastColLetter}${rowIndex}`;
    const amountRange = `${sheetName}!${amountCell}`;

    // Write label in a single cell only
    this.context.writeRange(labelRangeSingle, [[label]]);

    // Then merge the label area visually
    this.context.mergeRange(
      sheetName,
      `A${rowIndex}:${columnIndexToLetter(numCols - 1)}${rowIndex}`,
      "MERGE_ALL"
    );

    this.context.setAlignment(sheetName, fullRowRange, {
      horizontal: "RIGHT",
      vertical: "MIDDLE",
    });
    this.context.setBackground(sheetName, fullRowRange, "#e2e8f0");
    this.context.setTextFormat(sheetName, fullRowRange, {
      bold: true,
    });

    const dataStartRow = (this.props.dataStartRow);
    const dataEndRow = (this.props.dataEndRow);
    const amountColLetter = columnIndexToLetter(amountColumnIndex);
    const sumFormula = `=SUM(${amountColLetter}${dataStartRow}:${amountColLetter}${dataEndRow})`;

    this.context.writeRange(amountRange, [[sumFormula]]);
    this.context.setNumberFormat(sheetName, amountCell, "#,##0.00");
  }
}


/**
 * Root dashboard component that composes title, headers, data and summary.
 */
class DemoDashboardComponent extends SheetComponent {

  render() {
    const sheetName = (this.props.sheetName);
    const headers = (this.props.headers);
    const rows = (this.props.rows);

    this.renderChild(TitleComponent, {
      sheetName,
      title: "Mini Apps Script Component Demo",
      subtitle: "Rendered with a single batchUpdate + values.batchUpdate",
    });

    const headerRowIndex = 4;
    this.renderChild(HeaderRowComponent, {
      sheetName,
      headers,
      rowIndex: headerRowIndex,
    });

    const dataStartRow = headerRowIndex + 1;
    this.renderChild(DataTableComponent, {
      sheetName,
      rows,
      startRow: dataStartRow,
    });

    const dataEndRow = dataStartRow + rows.length - 1;
    const summaryRowIndex = dataEndRow + 2;

    this.renderChild(SummaryRowComponent, {
      sheetName,
      label: "Total amount",
      rowIndex: summaryRowIndex,
      numCols: headers.length,
      amountColumnIndex: headers.length,
      dataStartRow,
      dataEndRow,
    });

    const lastColLetter = columnIndexToLetter(headers.length);
    this.context.autoResizeColumns(sheetName, `A1:${lastColLetter}${summaryRowIndex}`);
  }
}


/**
 * Entry point: renders a full demo dashboard using the component system.
 *
 * @returns {void}
 */
function renderDemoDashboard() {
  const ss = SpreadsheetApp.openById("1mRno9TPj1iXGQ_lEqBIy3jOz4tyJm6upsVhklGuG1BQ")
  const spreadsheetId = ss.getId();
  const sheetName = "Demo";
  const headers = ["Project", "Owner", "Hours", "Amount"];
  const rows = [
    ["Surf Schedule App", "Lidia", 6, 750],
    ["Invoice Generator", "Lidia", 4, 520],
    ["ColorMyPie Add-on", "Lidia", 3, 380],
    ["Client Dashboard", "Lidia", 5, 640],
  ];

  renderSheet(spreadsheetId, DemoDashboardComponent, {
    sheetName,
    headers,
    rows,
  });
}
