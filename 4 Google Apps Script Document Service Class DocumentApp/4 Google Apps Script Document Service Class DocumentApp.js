
//Lesson 1 & 2

/* 
appendParagraph
getBody
DocumentApp.create
 */

function testDoc(){
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  body.appendParagraph('Hello World');
  body.appendPageBreak();
  Logger.log(body);
}


function addtoDoc(){
  const id = '1kDCr9jm1IovQ4GmpJYGDGToLmb6WZn6M679fF3Q6ZKU';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  body.appendParagraph('Hello World');
  body.appendHorizontalRule();
  body.appendPageBreak();
  Logger.log(doc);
}

function createmyDoc(){
  let myName = 'Tester Docs ';
  const doc = DocumentApp.create(myName);
  Logger.log(doc.getId());
  Logger.log(doc.getUrl());
  Logger.log(doc.getEditors());
  myName += ' ' + doc.getId();
  doc.setName(myName);
  const body = doc.getBody();
  body.appendParagraph('Hello World in '+myName);
  body.appendParagraph('URL ' + doc.getUrl());
  body.appendParagraph('NAME ' +doc.getName());
  body.appendParagraph('Editors ' +doc.getEditors());
  body.appendHorizontalRule();
  body.appendPageBreak();

}

//Lesson 3

/* 
getBody
getText
getParagraphs
loop : for
 */

function selContent(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const data = body.getText();
  const paras = body.getParagraphs();
  const spanVer = LanguageApp.translate(data,'en','es');
  body.appendHorizontalRule();
  body.appendParagraph('In Spanish');
  body.appendHorizontalRule();
  body.appendParagraph(spanVer);
  Logger.log(paras);

}


function updateContent(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  for(let i=0;i<10;i++){
    let temp = `${i} 
    'Hello' "Hi" More text being added .....
    `;
    body.appendParagraph(temp);
  }
}

//Lesson 4

/* 
getTextalignment
getType
loop : forEach
 */


function getParas(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const paras = body.getParagraphs();
  paras[0].setHeading(DocumentApp.ParagraphHeading.HEADING1);
  paras.forEach((p,index)=>{
    Logger.log(p.getTextAlignment());
    Logger.log(p.getType());
    let temp = p.getText();
    temp = index + '. '+p.getType()+' '+p.getTextAlignment()+' '+ temp;
    p.setText(temp);
    p.setTextAlignment(DocumentApp.TextAlignment.NORMAL);
  })
}

//Lesson 5

/* 
getNextSibling
getNumChildren [vaut zéro si pas de texte]
asText
 */

function getParas2(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const paras = body.getParagraphs();
  paras.forEach((p,index)=>{
    //Logger.log(p.getNumChildren());
    let temp = p.getText();
    //p.appendText(' ' + index);
  })
  Logger.log(paras[2].getText());
  Logger.log(paras[2].getNextSibling().asText().getText());
}

//Lesson 6 : ça commence à prendre du sens <<<ICI>>>

//getChild(i).getType()

/* 
getChild
asText
editAsTexst
 */

function bodyEle(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  //Logger.log(body.getNumChildren());
  for(let i=0;i<body.getNumChildren();i++){
    //Logger.log(body.getChild(i));
    let temp = body.getChild(i);
    //Logger.log(temp.getType());
    let contentInside = temp.asText().getText();
    //Logger.log(contentInside);
    if(temp.getType() == 'PARAGRAPH'){
      let val = temp.asParagraph().editAsText().setFontSize(20);
      if(temp.asParagraph().getText().length > 5){
        let endPos = 3;
        let startPos = 1;
        val.setBackgroundColor(startPos,endPos,'#e9c46a');
      }
      Logger.log(val);
      //temp.asParagraph().insertText(1,' NEW '+i); // provoque une erreur, parce que index à 1 pour un paragraphe vide provoque une erreur. Mais ça n'explique pas tout.
    }
  }
}

//Lesson 7 // edition du texte en continuité : body.editAsText().insertText

/* 
editAsText
insertText
 */

function docContents(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const txt = body.editAsText();
  Logger.log(txt.getText().length);
  txt.insertText(0,"#NEWCOLOR").setFontSize(0,8,30).setBackgroundColor(0,8,'#ff00ff').setForegroundColor(0,8,'#ffffff');
  txt.insertText(20,"#NEWBOLD").setBold(20,27,true);
  //val2.setBold(20,26,true);
  Logger.log(txt);

}

//Lesson 8

/* 
appendtext
getAttributes
 */

function addStyles(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const paras = body.getParagraphs();
  //Logger.log(paras.length);
  const style1 = {};
  style1[DocumentApp.Attribute.FONT_SIZE] = 22;
  style1[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ffffff';
  style1[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ff0000';
  const style2 = {};
  style2[DocumentApp.Attribute.FONT_SIZE] = 12;
  style2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  style2[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ffffff';
  paras.forEach((el,index)=>{
    if(el.getText().length > 0 ){
      Logger.log(el.getText().length);
      Logger.log(el.getNumChildren());
      Logger.log(el.getAttributes());
      let val = el.appendText('NEW');
      val.setAttributes(style1);
      if(index==2){
        el.setAttributes(style2);
      }
    }
  })
}

/*
{HORIZONTAL_ALIGNMENT=null, LINE_SPACING=null, HEADING=Normal, BOLD=null, LEFT_TO_RIGHT=true, BACKGROUND_COLOR=null, INDENT_END=null, INDENT_FIRST_LINE=null, STRIKETHROUGH=null, LINK_URL=null, FONT_FAMILY=null, INDENT_START=null, UNDERLINE=null, ITALIC=null, SPACING_AFTER=null, FONT_SIZE=12.0, SPACING_BEFORE=null, FOREGROUND_COLOR=null}
*/

//Lesson 9 recherche regex et remplacement

/* 
replaceText
findText
*/


function replacer(){
   const style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0000ff';
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ffff00';
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const body = DocumentApp.openById(id).getBody();
  const rep = body.replaceText('(?i)Lorem','LAURENCE');

 let founder = body.findText('LAURENCE');
  while (founder != null){
    let val = founder.getElement().asText();
    Logger.log(val);
    let start = founder.getStartOffset();
    let end = founder.getEndOffsetInclusive();
    Logger.log(start, end);
    
    val.setBackgroundColor(start,end,'#ff0000');
    val.setForegroundColor(start,end,'#ffffff');
    //val.setAttributes(style);
    founder = body.findText('LAURENCE',founder);
  }

  //Logger.log(founder);
  //founder.getElement().setAttributes(style);

  //rep.setAttributes(style);




}


Examples of regular expressions https://support.google.com/a/answer/1371417?hl=en
Regex Generator https://regexr.com/



//Lesson 10 gestion de liste

/* 
appendListItem
appendParagraph
getListId
*/

function addList(){
  const body = DocumentApp.getActiveDocument().getBody();
   //Logger.log(body.editAsText().getText());
  const val1 = body.appendListItem('item 1');

  Logger.log(val1.getListId());
//kix.uq8c9jsmuygr
  body.appendParagraph('new list');
  const val2 = body.appendListItem('item 2');
  val2.setListId(val1);
  Logger.log(val2.getListId());
  for(let i=0;i<10;i++){
    let val3 = body.appendListItem('item '+(i+2));
    val3.setListId(val1);
  }
}

//Lesson 11

// body.clear

// source 10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE
function addLorem(){
  const id = '10zw2TjR-6i5R6GKfUtFDWOUSBAlvqXeA8iu0oj39DLE';
  const body = DocumentApp.getActiveDocument().getBody();
  const sourceLorem = DocumentApp.openById(id).getBody();
  body.clear();
  const data = sourceLorem.getText();
  body.appendParagraph(data);
}



//Lesson 12

/* 
appendTable
insertTableCell
insertTableRow
*/

function addTable(){
  const style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0000ff';
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ffff00';
  const body = DocumentApp.getActiveDocument().getBody();
  body.clear();
  //const arr = [['col1','col2','col3'],['1','2','3'],['4','5','6']];
  const arr = getData();
  const val = body.appendTable(arr);
  Logger.log(val);
  const row = val.insertTableRow(4);
  row.insertTableCell(0,'TEST1');
  row.insertTableCell(1,'TEST2');
  row.insertTableCell(2,'TEST3');
  row.setAttributes(style);
}

function getData(){
  const id = '18EiSL1V4IAWvVi0vJsKc_JEIK1_bF4Ulo3L3Wb716II';
  const ss = SpreadsheetApp.openById(id).getSheets()[0];
  const data = ss.getDataRange().getValues();
  Logger.log(data);
  return data;
}



//Lesson 13

/* 
createMenu
addSeparator
addItem
addToUi
Button
ButtonSet
prompt
alert
*/

function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('Adv')
  .addItem('alert','popAlert')
  .addSeparator()
  .addItem('prompt','popPrompt')
  .addToUi();
}

function popAlert(){
  const ui = DocumentApp.getUi();
  const rep = ui.alert('Do you like Docs',ui.ButtonSet.YES_NO_CANCEL);
  let message = 'Why not - to bad';
  if(rep == ui.Button.YES){
    message = 'Great I\'m happy to hear that';
  }
  ui.alert(message);
}

function popPrompt(){
  const ui = DocumentApp.getUi();
  const rep = ui.prompt('Rate 1-5',ui.ButtonSet.OK);
  if(rep.getSelectedButton() == ui.Button.OK){
    let val = rep.getResponseText();
    ui.alert('We got the rating '+val);
  }else{
    ui.alert('Why did you close it???');
  }
}



//Lesson 14

// getCursor


function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('fun')
  .addItem('Lorem ipsum','adder')
  .addToUi();
}


function adder(){
    const cursor = DocumentApp.getActiveDocument().getCursor();
    const id = '1ma0ccRNdl9CSsbbdxDKkkY9sANzgi2S8Ylf7NTHK6Pk';
    const sourceContent = DocumentApp.openById(id).getBody().getText();

    if(cursor){
      const val = cursor.insertText(sourceContent);
      if(val){
        //val.setBackgroundColor('#ff0000');
        val.setBold(true);
      }
    }
}



//Lesson 15
/* 
getSelectedButton
getSelection
findText
*/

function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('fun')
  .addItem('Lorem ipsum','adder')
  .addItem('Highlighter','highlight')
  .addToUi();
}


function highlight(){
  const ui = DocumentApp.getUi();
  const body = DocumentApp.getActiveDocument().getBody();
  const rep = ui.prompt('Highlight What',ui.ButtonSet.OK_CANCEL);
  if(rep.getSelectedButton() == ui.Button.OK){
    let findThis = rep.getResponseText();
    let searchMe = body.findText(findThis);
    while(searchMe != null){
      let val = searchMe.getElement().asText();
      let start = searchMe.getStartOffset();
      let end = searchMe.getEndOffsetInclusive();
      val.setBackgroundColor(start,end,'#000000');
      val.setForegroundColor(start,end,'#000fff');
      searchMe = body.findText(findThis ,searchMe);
    }


  }

}




function adder(){
    const cursor = DocumentApp.getActiveDocument().getCursor();
    const id = '1ma0ccRNdl9CSsbbdxDKkkY9sANzgi2S8Ylf7NTHK6Pk';
    const sourceContent = DocumentApp.openById(id).getBody().getText();

    if(cursor){
      const val = cursor.insertText(sourceContent);
      if(val){
        //val.setBackgroundColor('#ff0000');
        val.setBold(true);
      }
    }
}



//Lesson 16

/* 
createHtmlOutput
createHtmlOutputFromFile
*/

function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('Adv')
  .addItem('html Modal','modal1')
  .addItem('file Modal','modal2')
  .addToUi();
}

function modal1(){
  const output = '<h1>Hello World</h1>';
  const html = HtmlService.createHtmlOutput(output)
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showModalDialog(html,'Title Popup');
}

function modal2(){
  const html = HtmlService.createHtmlOutputFromFile('popup')
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showModelessDialog(html,'Title Modal');
}

//Lesson 17

/* 
showSidebar
showModalDialog
showModelessDialog
*/

function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('Adv')
  .addItem('html Modal','modal1')
  .addItem('file Modal','modal2')
  .addItem('html Sidebar','side1')
  .addItem('file Sidebar','side2')
  .addToUi();
}

function side1(){
  const output = '<h1>Hello World</h1>';
  const html = HtmlService.createHtmlOutput(output)
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showSidebar(html);
}


function side2(){
  const html = HtmlService.createHtmlOutputFromFile('popup')
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showSidebar(html);
}


function modal1(){
  const output = '<h1>Hello World</h1>';
  const html = HtmlService.createHtmlOutput(output)
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showModalDialog(html,'Title Popup');
}

function modal2(){
  const html = HtmlService.createHtmlOutputFromFile('popup')
  .setWidth(600)
  .setHeight(500);
  DocumentApp.getUi().showModelessDialog(html,'Title Modal');
}



//Lesson 18 // à étudier de près, utiliser l'extension chrome Photos Direct Link pour que ça fonctionne !

/* 
UrlFetchApp
fetch
getBlob
insertImage
insertInlineImage
*/

function onOpen(){
  const ui = DocumentApp.getUi();
  ui.createMenu('Adv')
  .addItem('addImage','addImage')
  .addToUi();
}

function addImage(){
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const url = 'https://dummyimage.com/300x200/0fff00/f000ff';
  const myImage = UrlFetchApp.fetch(url).getBlob();
  cursor.insertInlineImage(myImage);

}


function insertImage(){
  const body = DocumentApp.getActiveDocument().getBody();
  const url = 'https://dummyimage.com/600x400/000/fff';
  const myImage = UrlFetchApp.fetch(url).getBlob();
  const img = body.insertImage(0,myImage);
  Logger.log(img);
  const img1 = body.appendImage(myImage);
}


