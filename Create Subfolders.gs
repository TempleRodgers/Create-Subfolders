/**
 * Temple Rodgers - 25/01/24
 * Simple script to take a list of folders
 * and create them as subfolders of the
 * folder where the script is running
 */
//when the sheet is opened create the menu to run the script 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script Menu')
      .addItem('Create Subfolders', 'createFolders')
      .addToUi();
}

function createFolders() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
    // find out about the current sheet and the folder it's in (above)
    // then check that a tab called 'Folders' is present,error if not
    var foldersSheet = spreadsheet.getSheetByName('Folders');
    if (!foldersSheet) {
      throw new Error('Sheet "Folders" not found. Please create it with the column "FolderName".');
    }
    // check the first row of the sheet to see that a column called
    // 'FolderName' is present. Error if not
    var headerRow = foldersSheet.getRange(1, 1, 1, foldersSheet.getLastColumn()).getValues()[0];
    var folderNameIndex = headerRow.indexOf('FolderName') + 1;
    if (folderNameIndex === 0) {
      throw new Error('Column "FolderName" not found in the "Folders" sheet.');
    }
    // pull all the folder names into a single variable folderNames
    var folderNames = foldersSheet.getRange(2, folderNameIndex, foldersSheet.getLastRow() - 1, 1).getValues().flat();
    // popup dialog box for user confirmation to create the folders
    var response = SpreadsheetApp.getUi().alert(
      'Create Subfolders',
      'Confirm folder creation in: ' + parentFolder.getName() + '\n\n(Move spreadsheet to change parent folder)',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    // get the foldernames one by one and create the folder
    if (response == SpreadsheetApp.getUi().Button.YES) {
      for (var i = 0; i < folderNames.length; i++) {
        var newFolderName = folderNames[i].trim(); // Trim whitespace
        if (newFolderName) { // Check for empty names
          var newFolder = parentFolder.createFolder(newFolderName); // create the folder
          Logger.log('Folder created: %s', newFolder.getName());
        }
      }
      SpreadsheetApp.getUi().alert('Folders created successfully!'); // Success message
    } else {
      Logger.log('Folder creation cancelled by user');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    Logger.log(error);
  }
}
