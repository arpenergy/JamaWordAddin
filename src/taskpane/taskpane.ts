/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
//const client = require('./jamaclient');


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("update-item").onclick = updateitem;
    document.getElementById("auto-format").onclick = autoformat;
  }
});

export async function autoformat() {
  return Word.run(async (context) => {
    //grab the document
    let doc = context.document;

    //To do - find document title style for title page detection
    //Step 2 - Delete Title page
    let myResults = doc.body.search("^m", { matchWildcards: true });
    myResults.load("items");
    await context.sync();

    let pagebreak = myResults.items[0].getRange("After");
    let secStart = doc.body.getRange("Start");
    let sectAll = secStart.expandTo(pagebreak);
    sectAll.delete();
    
    await context.sync();

    //Delete TOC - Todo
    myResults = doc.body.search("Table of Contents", { matchCase: true, matchWholeWord: true });
    myResults.load("items");
    await context.sync();
    let tocStart = myResults.items[0].getRange("Start");
    await context.sync();

    //Delete blank pages - disabled
    let paras = doc.body.paragraphs;
    paras.load("items");
    await context.sync();

    paras.items.forEach(function (para){
      if (para.text.trim().length === 0) {
        let selectedPara = para.getRange();
        //selectedPara.delete()
      }
    });

    //Step 3: Add "Document Information" Title
    let docInfo = paras.items[0].getRange().insertParagraph("Document Information","Before");
    docInfo.styleBuiltIn = Word.Style.heading1;
    await context.sync();

   
    //Step 4-6: Find and replace 
    myResults = doc.body.search("^m", { matchWildcards: true });
    if (!myResults) {
      myResults.load("items");
      await context.sync();
      myResults.items.forEach(function (item) {
        item.insertText(" ", "Replace");
      });
    }

    myResults = doc.body.search("^b", { matchWildcards: true });
    if (!myResults) {
      myResults.load("items");
      await context.sync();
      myResults.items.forEach(function (item) {
        item.insertText(" ", "Replace");
      });
    }

    myResults = doc.body.search("^t^p", { matchWildcards: true });
    if (!myResults) {
      myResults.load("items");
      await context.sync();
      myResults.items.forEach(function (item) {
        item.insertText("^p", "Replace");
      });
    }

    myResults = doc.body.search("^p^p^p", { matchWildcards: true });
    if (!myResults) {
      myResults.load("items");
      await context.sync();
      myResults.items.forEach(function (item) {
        item.insertText("^p^p", "Replace");
      });
    }





    //Step 7-8: Table formating
    let tbls = doc.body.tables;
    tbls.load("items");
    await context.sync();

    tbls.items.forEach(function (tbl) {
      tbl.autoFitWindow();
      tbl.styleBuiltIn = Word.Style.gridTable4_Accent6;
      tbl.styleBandedRows = true;
      
    });
    await context.sync();

    //Set all font to Calibre
    let bodyRange = doc.body.getRange();
    bodyRange.font.name = "Calibre";
    

    //Legacy requirements - R3000-MKT-00284-1.0:
    myResults = doc.body.search("R[0-9]*-*-[0-9]*-[0-9].[0-9]*:", { matchWildcards: true });
    myResults.load("items");
    await context.sync();

    myResults.items.forEach(function (item) {
      item.styleBuiltIn = Word.Style.heading3;
      item.insertText(item.text + "\r\n","Replace");
    });
    await context.sync();

  });
}


export async function updateitem() {
  return Word.run(async (context) => {
    //grab the document
    let doc = context.document;

    let selectedRange = context.document.getSelection();

    //grab the first hyperlink
    selectedRange.load(["items", "hyperlink"]);
    await context.sync();
    let link = selectedRange.hyperlink;

    if (link.includes("https://enphase.jamacloud.com/perspective.req?"))
    {
      let projectinfo = link.substring(link.indexOf('?') + 1).split("&");
      let projectId = parseInt(projectinfo[0].substring(10));
      let itemId = parseInt(projectinfo[1].substring(6));
      console.log("Project Id: ", projectId);
      console.log("Item Id: ", itemId);
    }


    //search for a particular word
    let myResults = selectedRange.search("Text Item ID: ", { matchCase: true, matchWholeWord: true });

    //load the properties
    myResults.load(["items", "text"]);
    await context.sync();

    //loop through the results
    myResults.items.forEach(function (rng) {
      console.log(rng.text);
    });

    

    //save it
    doc.save();
  });
}


