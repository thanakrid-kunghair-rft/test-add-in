/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("refresh-button").onclick = refreshButton;
  setOnSelectionChangedHandler();
};

async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function refreshButton() {
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.calculate();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function setOnSelectionChangedHandler() {
  try {
    await Excel.run(async context => {
      context.workbook.onSelectionChanged.add(onSelectionChangedHandler);
    });
  } catch (error) {
    console.error(error);
  }
}

async function onSelectionChangedHandler(args: Excel.SelectionChangedEventArgs) {
  await Excel.run(async (): Promise<void> => {
    const range = args.workbook.getSelectedRange();
    let directPrecedentValues: string[][][] = [];
    try {
      let spillParentRange = range.getSpillParentOrNullObject();
      range.load({ address: true, formulas: true });
      spillParentRange.load({ address: true, formulas: true });
      
      const precedentsRangeAreas: Excel.WorkbookRangeAreas = range.getDirectPrecedents();
      precedentsRangeAreas.ranges.load('values');
      await args.workbook.context.sync();
      if (precedentsRangeAreas.ranges) {
        directPrecedentValues = precedentsRangeAreas.ranges.items.map((item: Excel.Range) => item.values);
      }
      console.log('DirecPrecedent Values')
      console.log(directPrecedentValues)
    } catch(e) {
      console.error(e)
    }
    
  });
}


