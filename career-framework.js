var fontFaceName = 'Open Sans';
var email = 'jeremy.burns@news.co.uk'
var templateFolderId = '115Umgt4Mlfey1YwBgtU6Vbgh-c3EUG7u';
var firstCompetencySheet = 1;

function onOpen() {
  buildMenu()
  addRolesToOverviewSheet();
}

function buildMenu() {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Career Framework')
  
  menu
    .addItem('Build Template', 'buildTemplate')
    .addItem('Build All Templates', 'buildAllTemplates')
    // .addItem('Get Roles', 'getRoles')
    .addToUi();

  var roles = getRoles()
  addRolesToOverviewSheet(roles)
 }

function getRoles() {
  var sheetRoles = SpreadsheetApp.getActiveSpreadsheet().getSheets()[this.firstCompetencySheet];
  var range = sheetRoles.getDataRange();
  var rows = range.getValues();
  var firstColumn = 2;
  var firstRow = rows[0];
  var roles = [];
  
  for (var col = firstColumn; col < firstRow.length; col++) {
    roles.push(firstRow[col]);
  }

  return roles;
}

function addRolesToOverviewSheet(roles) {

  if (!roles.length)
    roles = getRoles()

  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetOverview = sheet.getSheetByName('Overview');

  sheet.getRangeByName('Overview!job_roles').clearContent()

  row = sheet.getLastRow();

  for (var roleNumber = 0; roleNumber < roles.length; roleNumber++) {
    sheetOverview.appendRow([roles[roleNumber]])
  }
}

function getRoleNumber(roleName) {
  var roles = getRoles();
  
  for (var roleNumber = 0; roleNumber < roles.length; roleNumber++) {
    if (roles[roleNumber]  === roleName)
      return roleNumber;
  }

  return false;
}

function buildAllTemplates() {
  var roles = getRoles();
  var roleName;

  for (var roleNumber = 0; roleNumber < roles.length; roleNumber++) {
    roleName = roles[roleNumber];
    if (!build(roleName)) {
      Browser.msgBox('Template build failed', 'Failed to build template for ' + roleName, Browser.Buttons.OK);
      return;
    }
  }
  
  Browser.msgBox('Template building complete', 'All templates built.', Browser.Buttons.OK);
  
}

function buildTemplate() {
  
  var roleName = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  
  if (build(roleName)) {
    Browser.msgBox('Template build complete', 'Building template for ' + roleName + ' completed.', Browser.Buttons.OK);
  } else {
    Browser.msgBox('Template build failed', 'Failed to build template for ' + roleName, Browser.Buttons.OK);
  }
  
}

function build(roleName) {
  var roleNumber = getRoleNumber(roleName);
  
  if (roleNumber === false) {
    Browser.msgBox('Invalid role', 'Please select a cell that contains a valid role name.', Browser.Buttons.OK);
    return false;
  }
  var templateSpreadsheet = createTemplate(roleName);
  var firstCompetencySheet = this.firstCompetencySheet;
  var competencies = getCompetencies(roleNumber, firstCompetencySheet);
  
  if (competencies == []) {
    Browser.msgBox('Template build failed', 'No competencies available for ' + roleName, Browser.Buttons.OK);
    return false;
  }
  
  fillDetails(templateSpreadsheet, roleName)
  fillScores(templateSpreadsheet, competencies)

  var fileId;
  try {
    fileId = templateSpreadsheet.getId();
  }
  catch (err) {
    Browser.msgBox('Invalid fileid', err, Browser.Buttons.OK);
  }

  try {
    moveToTemplateFolder(fileId, this.templateFolderId)
  }
  catch (err) {
    Browser.msgBox('Failed template move', err, Browser.Buttons.OK);
  }

  return true;

}

function createTemplate(roleName) {
  var templateSpreadsheet = SpreadsheetApp.create(roleName);
  templateSpreadsheet.renameActiveSheet('Details');
  templateSpreadsheet.insertSheet('Overview');
  templateSpreadsheet.insertSheet('Scores');

  return templateSpreadsheet;
}

function getCompetencies(roleNumber, firstCompetencySheet) {
  var competencies = {};

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (var sheetNum = firstCompetencySheet; sheetNum < sheets.length; sheetNum++) {    

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetNum];
    var competencyName = sheet.getName();
    competencies[competencyName] = [];
    var rangeData = sheet.getDataRange();
    var lastRow = sheet.getLastRow();
    var scoreCol = roleNumber + 2;
    var rangeData = sheet.getRange(2, 1, lastRow, scoreCol + 1)
    var skills = rangeData.getValues();

    for (var skillNum = 0; skillNum < skills.length; skillNum++) {
      var skill = skills[skillNum];

      if (skill[0] == '')
        break;

      if (!isNaN(parseFloat(skill[scoreCol])) && isFinite(skill[scoreCol])) {
        competencies[competencyName].push([
          skill[0],
          skill[1],
          skill[scoreCol]
        ]);
      }
    }
  }

  return competencies;
}

function addOverviewChart(templateSpreadsheet, rangeChartTopLeft, rangeChartBottomRight) {
  var sheetOverview = templateSpreadsheet.getSheetByName('Overview');

  var summaryChart = sheetOverview.newChart()
     .setChartType(Charts.ChartType.COLUMN)
     .setOption('width', 750)
     .setOption('height', 500)
     .setOption('legend', 'none')
     .setPosition(2, 4, 0, 0)
     .addRange(sheetOverview.getRange(rangeChartTopLeft + ':' + rangeChartBottomRight))
     .setOption('vAxes', {
       0: {viewWindow: {min: 0, max: 1000}}})
     .build();
  
     sheetOverview.insertChart(summaryChart);
  
}

function fillDetails(templateSpreadsheet, roleName) {
  var sheetDetails = templateSpreadsheet.getSheetByName('Details');
  
  sheetDetails
    .appendRow(['Engineering Career Framework'])
    .appendRow([roleName])
    .appendRow([''])
    .appendRow(['Name:'])
    .appendRow(['Area:'])
    .appendRow(['Role:'])
    .appendRow(['Completed with:'])
    .appendRow(['Date completed:'])
    .appendRow(['Email address:'])
    .appendRow(['Manager:'])
    .appendRow(['Github handle:']);
  
  sheetDetails.getRange('A:B')
    .setFontSize(11)
    .setFontFamily(this.fontFaceName);

  sheetDetails
    .getRange('A2')
    .setFontSize(12);
  
  sheetDetails
    .getRange('A:A')
    .setFontWeight('bold');
    
  sheetDetails
    .autoResizeColumn(1)
    .setColumnWidth(2, 200);
}

function fillScores(templateSpreadsheet, competencies) {
  var sheetScores = templateSpreadsheet.getSheetByName('Scores');
  var sheetOverview = templateSpreadsheet.getSheetByName('Overview');
  var rangeCompetencyName, rangeCompetencyFirstRow, rangeCompetencyLastRow, rangeChartTopLeft, rangeChartBottomRight;
  var competencyCount = 1;

  sheetOverview.appendRow(['Competency', 'Score']);
  sheetOverview.getRange('A1:B1')
      .setFontSize(12)
      .setFontWeight('bold');

  Object.keys(competencies).forEach(function(competencyName) {

    competencyCount++;
    
    sheetScores.appendRow([competencyName,'Description', 'Weighting', 'Score', 'Total', 'Comments']);

    thisRow = sheetScores.getLastRow();

    rangeCompetencyName = thisRow;
    
    sheetScores.getRange('A' + thisRow + ':F' + thisRow)
      .setFontSize(10)
      .setFontWeight('bold')
      .setBackground('#efefef');
      
    sheetScores.getRange(thisRow, 1).setFontSize(12);

    if (!rangeChartTopLeft)
      rangeChartTopLeft = 'A' + competencyCount;

    skills = competencies[competencyName];

    skills.forEach(function(skill) {

      sheetScores.appendRow([skill[0], skill[1], skill[2]]);

      thisRow = sheetScores.getLastRow();

      if (!rangeCompetencyFirstRow)
        rangeCompetencyFirstRow = thisRow;
      
      sheetScores.getRange(thisRow, 5)
        .setFormula('=C' + thisRow + '*D' + thisRow);
      
      sheetScores.getRange(thisRow, 4)
        .setDataValidation(SpreadsheetApp.newDataValidation()
          .setAllowInvalid(false)
          .setHelpText('Please choose 0, 0.5 or 1')
          .requireNumberBetween(0, 1)
          .build());
    });

    rangeCompetencyLastRow = thisRow;

    sheetScores.appendRow(['','Subtotal:']);

    thisRow = sheetScores.getLastRow();

    sheetScores.getRange(thisRow, 3)
      .setFormula('=SUM(C' + rangeCompetencyFirstRow + ':C' + rangeCompetencyLastRow + ')');
    sheetScores.getRange(thisRow, 4)
      .setFormula('=SUM(D' + rangeCompetencyFirstRow + ':D' + rangeCompetencyLastRow + ')');
    sheetScores.getRange(thisRow, 5)
      .setFormula('=SUM(E' + rangeCompetencyFirstRow + ':E' + rangeCompetencyLastRow + ')');
    sheetScores.getRange('B' + thisRow + ':E' + thisRow)
      .setFontSize(10)
      .setFontWeight('bold');

    sheetOverview.appendRow([
      '=Scores!A' + rangeCompetencyName,
      '=Scores!E' + thisRow,
    ]);
    
    sheetScores.appendRow([' ']);

    rangeCompetencyFirstRow = 0;

  });

  rangeChartBottomRight = 'B' + competencyCount;

  sheetScores
    .setColumnWidth(1, 250)
    .setColumnWidth(2, 250)
    .setColumnWidth(3, 50)
    .setColumnWidth(4, 50)
    .setColumnWidth(5, 50)
    .setColumnWidth(6, 400)

  sheetScores.getDataRange()
    .setFontFamily(this.fontFaceName);

  sheetScores.getRange('A:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheetScores.getDataRange().setVerticalAlignment('top');

  sheetOverview.getDataRange()
    .setFontSize(12)
    .setFontFamily(this.fontFaceName);
  
  sheetOverview
    .autoResizeColumn(1)
    .autoResizeColumn(2)

  var email = this.email;

  protectRangebyEmail(sheetScores, 'A:A', 'Lock down the skills', email);
  protectRangebyEmail(sheetScores, 'B:B', 'Lock down the descriptions', email);
  protectRangebyEmail(sheetScores, 'C:C', 'Lock down the weightings', email);
  protectRangebyEmail(sheetScores, 'E:E', 'Lock down the totals', email);

  addOverviewChart(templateSpreadsheet, rangeChartTopLeft, rangeChartBottomRight);
}

function protectRangebyEmail(sheet, range, description, email) {
  var protection = sheet
    .getRange(range)
    .protect()
    .setDescription(description);
  
  protection.addEditor(email);
  
  protection.removeEditors(protection.getEditors());
  
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

}

function moveToTemplateFolder(fileID, targetFolderID) {
  var file
  try {
    file = DriveApp.getFileById(fileID);
  }
  catch (err) {
    Browser.msgBox('Error getting file', err, Browser.Buttons.OK);
  }
  var parents = file.getParents();

  while (parents.hasNext()) {
    var parent = parents.next();
    try {
      parent.removeFile(file);
    }
    catch (err) {
      Browser.msgBox('Error getting parents', err, Browser.Buttons.OK);
    }
    
  }

  try {
    DriveApp.getFolderById(targetFolderID).addFile(file);
  }
  catch(err) {
    Browser.msgBox('Error moving file', err, Browser.Buttons.OK);
  } 
  
}
