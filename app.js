//
// 1. Characters that aren't allowed in SharePoint Online:
//                " * : < > ? / \ |
// 2. "_vti_" cannot appear anywhere in a file or folder name
// 3. The entire path, including the file name, must contain fewer than 400 characters
// 4. You can’t create a folder name in SharePoint Online that begins with a tilde (~).
// 5. 15GB File upload size limit
//

const XLSX = require('xlsx');
const appRoot = require('app-root-path');
const winston = require('winston');

const { combine, timestamp, printf } = winston.format;
const myFormat = printf(({ timestamp, level, message, meta }) => {
  return `${timestamp} : ${level} : ${message}`;
});
const logger = winston.createLogger({
    format: combine(
      timestamp(),
      myFormat
    ),
    transports: [
        //new winston.transports.Console(),
        new winston.transports.File({ filename: `${appRoot}/log/app.log` })
    ]
});

//Open the Excel file
logger.info('Opening file...');
var workbook = XLSX.readFile(`${appRoot}/data/folder_tree_run_on_4-2-19__11-04-31-PM-Sheet1.xlsx`);
logger.info('Opening file...Complete!');

//Loop through the three sheets
//for(var sheets = 0; sheets <= 2; sheets++){
//  var sheet = workbook.Sheets[workbook.SheetNames[sheets]];

var sheet = workbook.Sheets[workbook.SheetNames[0]];
var range = XLSX.utils.decode_range(sheet['!ref']);
//Loop through each row
for(var rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
  var ownerLogin = sheet[XLSX.utils.encode_cell({r: rowNum, c: 1})];
  var path = sheet[XLSX.utils.encode_cell({r: rowNum, c: 2})];
  var item = sheet[XLSX.utils.encode_cell({r: rowNum, c: 4})];
  var id = sheet[XLSX.utils.encode_cell({r: rowNum, c: 5})];
  var type = sheet[XLSX.utils.encode_cell({r: rowNum, c: 6})];
  var size = sheet[XLSX.utils.encode_cell({r: rowNum, c: 7})];

  //Check for invalid characters
  if(item.v.includes('\"')){
    logger.warn('Invalid character \": ' + ownerLogin.v + item.v + id.v);
    //replace with ' single quote
  }
  if(item.v.includes('?') || item.v.includes('*')){
    logger.warn('Invalid character [? *]: ' + ownerLogin.v + item.v + id.v);
    //console.log('replace with . period');
  }
  if(item.v.includes(':') || item.v.includes('|')){
    logger.warn('Invalid character [: |]: ' + ownerLogin.v + item.v + id.v);
    //console.log('replace with - hyphen');
  }
  if(item.v.includes('<') || item.v.includes('>') || item.v.includes('/') || item.v.includes('\\')){
    logger.warn('Invalid character [< > / \\]: ' + ownerLogin.v + item.v + id.v);
    //replace with nothing
  }
  //"_vti_" cannot appear anywhere in a file or folder name
  if(item.v.includes('_vti_')){
    logger.warn('Invalid filename _vti_: ' + ownerLogin.v + item.v + id.v);
    //replace with vti-removed-in-migration single quote
  }
  // The entire path, including the file name, must contain fewer than 400 characters
  if(path.v.length > 400){
    logger.warn('Path name is longer than 400 chars: ' + ownerLogin.v + path.v + id.v);
    //what do we do?
  }
  //You can’t have folder names in SharePoint Online that begins with a tilde (~).
  if(item.v.startsWith('~') & type.v.includes('Folder')){
    logger.warn('Invalid folder starting with ~: ' + ownerLogin.v + item.v + id.v);
    //replace with tilde-removed-in-migration single quote
  }
  //15GB File upload size limit
  if(size.v.includes('GB') & type.v.includes('File')){
    var NUMERIC_REGEXP = /[-]{0,1}[\d]*[\.]{0,1}[\d]+/g;
    var gigabytes = parseFloat(Number(size.v.match(NUMERIC_REGEXP)));
    if(gigabytes > 15.0) {
      logger.warn('File larger than 15GB: ' + ownerLogin.v + item.v + id.v + size.v);
      //what do we do?
    }
  }
}
