/*
*Chat Markup Formatter*
by Eden Biskin (SightSpirit)
https://github.com/SightSpirit/
https://www.ultimalia.net

[bold(), italic(), and strike() are adapted from https://stackoverflow.com/a/42929145]
Distributed under the MIT License. Please don't use this code for commercial purposes.

Interprets and removes formatting markup from chat apps like Slack and Rocket.Chat. Supports bold, italics, and strikethrough.

Install this script as a Google Docs Add-on: [URL coming soon; currently migrating to Google Workspace Marketplace]
*/

function bold() {
  var body = DocumentApp.getActiveDocument().getBody();
  var foundElement = body.findText("\\*.+\\*");
  
  while (foundElement != null) {
    var foundText = foundElement.getElement().asText();

    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    
    foundText.setBold(start,end,true);
    
    foundText.deleteText(start, start);
    foundText.deleteText(end-1, end-1);
    
    foundElement = body.findText("\\*.+\\*", foundElement);
   }
}

function italic() {
  var body = DocumentApp.getActiveDocument().getBody();
  var foundElement = body.findText("\\_.+\\_");
  
  while (foundElement != null) {
    var foundText = foundElement.getElement().asText();

    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    
    foundText.setItalic(start,end,true);
    
    foundText.deleteText(start, start);
    foundText.deleteText(end-1, end-1);
    
    foundElement = body.findText("\\_.+\\_", foundElement);
   }
}

function strike() {
  var body = DocumentApp.getActiveDocument().getBody();
  var foundElement = body.findText("\\~.+\\~");
  
  while (foundElement != null) {
    var foundText = foundElement.getElement().asText();

    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    
    foundText.setStrikethrough(start,end,true);
    
    foundText.deleteText(start, start);
    foundText.deleteText(end-1, end-1);
    
    foundElement = body.findText("\\~.+\\~", foundElement);
   }
}

function run() {
  this.bold();
  this.italic();
  this.strike();
}

function onOpen(e) {
  var menu = DocumentApp.getUi().createMenu("Chat Markup")
  menu.addItem("Run!", "run");
  menu.addToUi();
}
