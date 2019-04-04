//
// 1. Characters that aren't allowed in SharePoint Online:
//                " * : < > ? / \ |
// 2. "_vti_" cannot appear anywhere in a file or folder name
// 3. The entire path, including the file name, must contain fewer than 400 characters
// 4. You can’t create a folder name in SharePoint Online that begins with a tilde (~).
// 5. 15GB File upload size limit
//
const config = require('./config');
const XLSX = require('xlsx');
const appRoot = require('app-root-path');
const winston = require('winston');
const BoxSDK = require('box-node-sdk');

// Initialise the logger
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

// Initialise the Box SDK with app credentials
var sdk = new BoxSDK({
  clientID: config.boxAuth.clientID,
  clientSecret: config.boxAuth.clientSecret
});

// Create a basic API client, which does not automatically refresh the access token
var client = sdk.getBasicClient(config.boxAuth.developerToken);

//Open the Excel file
logger.info('Opening file: ' + `${appRoot}/data/${config.filename}`);
var workbook = XLSX.readFile(`${appRoot}/data/${config.filename}`);


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
  if(item.v.includes('?') ||
    item.v.includes('*') ||
    item.v.includes(':') ||
    item.v.includes('|') ||
    item.v.includes('\"') ||
    item.v.includes('<') ||
    item.v.includes('>') ||
    item.v.includes('/') ||
    item.v.includes('\\') ||
    item.v.includes('_vti_')
    ){
      logger.warn('Invalid character: ' + id.v + ' ' + type.v + ' ' + item.v);

      //Replace with permitted character
      var newName = item.v.replace('?', ' ');
      newName = newName.replace('*', ' ');
      newName = newName.replace(':', '-');
      newName = newName.replace('|', '-');
      newName = newName.replace('\"', '\'');
      newName = newName.replace('<', '');
      newName = newName.replace('>', '');
      newName = newName.replace('/', '');
      newName = newName.replace('\\', '');
      newName = newName.replace('_vti_', 'vti-removed-in-migration');

      if(ownerLogin.v.includes('boxadmin@krb.nsw.edu.au')){
        if(type.v.includes('File')){
          updateBoxFile(id.v, newName);
        } else if (type.v.includes('Folder')){
          updateBoxFolder(id.v, newName);
        } else {
          logger.error('Unknown object type: ' + id.v);
        }
      } else {
        logger.error('Object not owned by boxadmin and needs a new name: ' + id.v + ' ' + ownerLogin.v);
      }

  }

  //You can’t have folder names in SharePoint Online that begins with a tilde (~).
  if(item.v.startsWith('~') & type.v.includes('Folder')){
    logger.warn('Invalid folder starting with ~: ' + id.v + ' ' + item.v);
    var newName = item.v.replace('~', 'migrated-');
    //replace tilde with 'migrated-...'
    if(ownerLogin.v.includes('boxadmin@krb.nsw.edu.au')){
      updateBoxFolder(id.v, newName);
    } else {
      logger.error('Object not owned by boxadmin and needs a new name: ' + id.v + ' ' + ownerLogin.v);
    }
  }

  // The entire path, including the file name, must contain fewer than 400 characters
  if(path.v.length > 400){
    logger.error('Path name is longer than 400 chars: ' + id.v + ' ' + path.v );
    //what do we do?
  }

  //15GB File upload size limit
  if(size.v.includes('GB') & type.v.includes('File')){
    var NUMERIC_REGEXP = /[-]{0,1}[\d]*[\.]{0,1}[\d]+/g;
    var gigabytes = parseFloat(Number(size.v.match(NUMERIC_REGEXP)));
    if(gigabytes > 15.0) {
      logger.error('File larger than 15GB: ' + id.v + ' ' + ownerLogin.v + item.v + size.v);
      //what do we do?
    }
  }
}

// Call Box SDK API
function updateBoxFile(boxID, newFileName){
    client.files.get(boxID)
      .then(file => {
        //logger.info('Box file to rename: ' + boxID + ' ' + file.name);
        logger.info('client.files.update: ' + boxID + ' ' + newFileName);
      })
      .catch(err => logger.error('Box file API GET error: ' + boxID + ' ' + err));;
}
function updateBoxFolder(boxID, newFolderName){
  client.folders.get(boxID)
  	.then(folder => {
      //logger.info('Box folder to rename: ' +  folder.name);
      logger.info('client.folders.update: ' + boxID + ' ' + newFolderName);
  	})
    .catch(err => logger.error('Box folder API GET error: ' + boxID + ' ' + err));
}
