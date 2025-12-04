<h1>RenderSheet — A React-Inspired Renderer for Google Sheets (Apps Script)</h1>

<p align="left">
  <a href="LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="MIT License">
  </a>
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-Compatible-34A853?logo=google" alt="Apps Script">
  <img src="https://img.shields.io/badge/Requires-Google%20Sheets%20API-yellow?logo=google" alt="Sheets API">
  <img src="https://img.shields.io/badge/Project-Active-success" alt="Status">
</p>

RenderSheet is a lightweight, component-based rendering engine for **Google Sheets**.  
Inspired by **React**, it lets you build dashboards, tables, and spreadsheet layouts using **components**, **props**, and a single rendering pass powered by the `batchUpdate` Sheets API.

Fast, maintainable, and ideal for Apps Script developers.

---

## Features

- Automatic **A1 → GridRange** conversion  
- Formatting helpers:
  - Backgrounds  
  - Text formatting  
  - Borders  
  - Alignment  
  - Data validation  
  - Merge / unmerge  
  - Wrap strategy  
  - Auto-resize columns  
- Zero repeated API calls → **extremely fast**  
- Pure Apps Script — no dependencies  

---

## Requirements

RenderSheet requires the **Advanced Google Sheets API**.

Enable it here:
Apps Script → Services → Google Sheets API → Add


---

## Concept Overview

RenderSheet behaves like a mini React engine for Sheets:

---

### ✔️ 1. You write components

```js
class MyComponent extends SheetComponent {
  render() {
    this.context.writeRange("Sheet1!A1", [["Hello"]]);
    this.context.setBackground("Sheet1", "A1", "#000000");
  }
}
```
✔️ 2. You compose them
```js
class Dashboard extends SheetComponent {
  render() {
    this.renderChild(HeaderComponent, { title: "My Dashboard" });
    this.renderChild(DataTableComponent, { rows: myData });
  }
}
```
✔️ 3. You render once
```js
renderSheet(ssId, Dashboard, {
  sheetName: "Demo",
  rows: myData
});
```

RenderSheet collects everything and executes:

1 formatting batchUpdate
1 values batchUpdate

-> Two API calls total.


## Creating Your Own Components

Create reusable Sheets UI blocks in seconds.

Example: Status Pill Component
```js
class StatusPill extends SheetComponent {
  render() {
    const { sheetName, range, status } = this.props;

    const color =
      status === "Done" ? "#10b981" :
      status === "Blocked" ? "#ef4444" :
      "#facc15";

    this.context.setBackground(sheetName, range, color);
    this.context.setTextFormat(sheetName, range, {
      bold: true,
      color: "#ffffff"
    });
    this.context.setAlignment(sheetName, range, { horizontal: "CENTER" });
  }
}
```

Use it like:
```js
this.renderChild(StatusPill, {
  sheetName: "Demo",
  range: "E5",
  status: "Done"
});
```


