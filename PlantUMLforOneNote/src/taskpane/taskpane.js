/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your OneNote code here
   */
  try {
    await OneNote.run(async (context) => {

        // Get the current page.
        const pageContents = context.application.getActivePage().contents;        
        pageContents.load("items");        
        
        await context.sync();

        // Iterate through the page contents.
        for (let i = 0; i < pageContents.items.length; i++) {
            const content = pageContents.items[i];            
            console.log(content.id, " ", content.type, " ", content.title);
            
            if (content.type === "Outline") {                            
              const paragraphs = content.outline.paragraphs;
              paragraphs.load("items");
                
              await context.sync();

              //iterate through the paragraphs
              for (let j = 0; j < paragraphs.items.length; j++) {
                  const paragraph = paragraphs.items[j];                    
                  console.log("Paragraph:", paragraph.id, " ", paragraph.type);
                  if (paragraph.type === "RichText") {
                    paragraph.richText.load();
                    await context.sync();
                    console.log("Rich Text:", paragraph.id, " ", paragraph.type, " ", paragraph.richText.text, " ", paragraph.richText.style, " ",paragraph.richText.getHtml());
                  }                  
                  else if (paragraph.type === "Image") {
                    paragraph.image.load();
                    await context.sync();
                    console.log("Image:    ", paragraph.id, " ", paragraph.type, " ", paragraph.image.url);
                  }  
                  else if (paragraph.type === "Table") {
                    paragraph.table.load();
                    await context.sync();
                    console.log("Table:    ", paragraph.id, " ", paragraph.type, " ", paragraph.table.rowCount, " ", paragraph.table.columnCount);
                  }
                  else
                  {
                    console.log("Other:    ", paragraph.id, " ", paragraph.type);
                  }                                              
              }
            }
        }

                
        // Run the queued commands, and return a promise to indicate task completion.
        
    });
} catch (error) {
    console.log("Error: " + error);
}
}
