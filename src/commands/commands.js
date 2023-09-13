/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  Excel.run(async context => {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);  
    let sheets = context.workbook.worksheets;
    sheets.onActivated.add(onActivate);
    await context.sync();
    console.log("A handler has been registered for the onActivated sheets event. ");
    Event.completed();
  });  
  console.log("commands js onready run");

});

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */

async function writeValue(event) {
  
}

async function onActivate(event) {
  await Excel.run(async (context) => {    
      
      let sheet = context.workbook.worksheets.getActiveWorksheet();      
      sheet.load("name");
      await context.sync();           
      sheet.onSingleClicked.add(onSingleClick);       
      console.log(`Sheet activated: ${sheet.name}` );      
});
}

async function onSingleClick(event) {
  await Excel.run(async (context) => {    
      
    let range = context.workbook.getSelectedRange();
    range.load("address");

    await context.sync();
    
    console.log(`The address of the selected range is "${range.address}"`);
      
});
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

Office.actions.associate("writeValue", writeValue);