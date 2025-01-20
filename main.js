const presentation = SlidesApp.getActivePresentation();
var folder_name = "PicturesToImport" // folder in google drive name
var sortBy = '' // made this global var because im lazy (the inputting vars for these function a little freaky)
var format = '' // ^^^^^^^^

function onOpen() {
  const ui = SlidesApp.getUi();

  const sortingSubmenu = ui.createMenu("Sorting Options")
    .addItem("Sort Alphabetically", "setAlphabeticalSort")
    .addItem("Sort by Natural Order", "setNaturalOrderSort")
    //.addItem("Sort by Date Created", "runDateSort");

  const formatSubmenu = ui.createMenu("Format Options")
    .addItem("No Text", "setFormatBlank")
    .addItem("File Name As Title", "setFormatTitle")
    //.addItem("Sort by Date Created", "runDateSort");
    
   ui.createMenu("Mass Image Adding")
  .addSubMenu(sortingSubmenu)
  .addSubMenu(formatSubmenu)
  .addItem('Add Images', "main")
  .addToUi();

}


function getMyFiles(sortBy = "natural") {
  const folder = DriveApp.getFoldersByName("PicturesToImport").next();
  const files = folder.getFiles();
  const fileInfo = [];

  sortBy = PropertiesService.getUserProperties().getProperty('sortBy') || sortBy; // retreive stored var

  while (files.hasNext()) {
    const file = files.next();
    fileInfo.push({
      id: file.getId(),
      name: file.getName(),
      dateCreated: file.getDateCreated()
    });
  }

  if (sortBy === "alphabetical") {
    // Sort by name
    fileInfo.sort((a, b) => a.name.localeCompare(b.name));
  } else {
    // Sort by creation date
    fileInfo.sort((a, b) => a.dateCreated - b.dateCreated);
  }

  Logger.log("Files found: " + fileInfo.length);
  return fileInfo;
}


 

function addImageSlide(fileInfo) {

  format = PropertiesService.getUserProperties().getProperty('format') || 'blank'; // retrieve stored var
  if (format == 'title'){
    
    slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    
    // find title in all elements in slide
    const shapes = slide.getShapes();
    for (let shape of shapes) {
      shape.getText().setText(fileInfo.name);  // set the title text
      break;
      }
      
    }
  else{
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  }
  

  var image = slide.insertImage(DriveApp.getFileById(fileInfo.id));
  //SlidesApp.getUi().alert(fileInfo.name); //alert out each file name as they are added
  var imgWidth = image.getWidth();

  var imgHeight = image.getHeight();

  var pageWidth = presentation.getPageWidth();

  var pageHeight = presentation.getPageHeight();

  var newX = pageWidth/2 - imgWidth/2;

  var newY = pageHeight/2 - imgHeight/2;

  image.setLeft(newX).setTop(newY);

  // put image back on z index so title goes infront
  image.sendBackward();

}

// set sorting & format based on ui
function setAlphabeticalSort() {
  sortBy = 'alphabetical';
  PropertiesService.getUserProperties().setProperty('sortBy', sortBy); // save to user props so isn't reset everytime (appscript runs whole script every time any part of the script is run, which includes global var defenitions, which is very dumb and unintuitive)
}

function setNaturalOrderSort() {
  sortBy = 'natural';
  PropertiesService.getUserProperties().setProperty('sortBy', sortBy); // ^
}

function setFormatBlank() {
  format = 'blank';
  PropertiesService.getUserProperties().setProperty('format', format); // ^
}

function setFormatTitle() {
  format = 'title';
  PropertiesService.getUserProperties().setProperty('format', format); // ^
}

function main() {
  if (!presentation) {
    Logger.log("No active presentation found. Please open a presentation first.");
    return;
  }

  const fileInfoArray = getMyFiles(sortBy);
  fileInfoArray.forEach(addImageSlide);

  Logger.log("Slides appended to presentation: " + presentation.getUrl());
}
