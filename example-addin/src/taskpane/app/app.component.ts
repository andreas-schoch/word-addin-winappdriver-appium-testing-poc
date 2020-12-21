import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global require, Word */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  }
}



// Run button xPath by WAD UI Recorder
// "/Pane[@ClassName=#32769][@Name=Desktop 1]"
// "/Window[@ClassName=OpusApp][@Name=Document2 - Word]"
// "/Pane[@ClassName=MsoCommandBarDock][@Name=MsoDockRight]"
// "/ToolBar[@ClassName=MsoCommandBar]"
// "/Pane[@ClassName=MsoWorkPane][@Name=My Office Add-in]"
// "/Pane[@ClassName=NUIPane]"
// "/Pane[@ClassName=NetUIHWNDElement]"
// "/Custom[@ClassName=NetUInetpane][@Name=My Office Add-in]"
// "/Custom[@ClassName=NetUIOcxControl]"
// "/Pane[@ClassName=OsfAxControl]"
// "/Pane[@ClassName=Win32WebViewHolder]"
// "/Pane[@ClassName=CoreApplicationBridgeWindow]"
// "/Window[@ClassName=Windows.UI.Core.CoreWindow]"
// "/Window[@ClassName=Windows.UI.Core.CoreWindow]"
// "/Pane[@ClassName=Internet Explorer_Server][@Name=My Office Add-in]"
// "/Pane[@Name=My Office Add-in]"
// "/Group[position()=2]"
// "/Button[@Name=Run]"
