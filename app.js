//
// 1. Characters that aren't allowed in SharePoint Online:
//                " * : < > ? / \ |
// 2. "_vti_" cannot appear anywhere in a file or folder name
// 3. The entire path, including the file name, must contain fewer than 400 characters
// 4. You can’t create a folder name in SharePoint Online that begins with a tilde (~).
// 5. 15GB File upload size limit
//
const config = require('./config');
const csv = require('csv-parser');
const fs = require('fs');
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
        new winston.transports.Console({ level: 'error'}),
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
fs.createReadStream(`${appRoot}/data/${config.filename}`)
  .pipe(csv())
  .on('data', (row) => {
    //Loop through each row

    //Ignore anything in CareyG's home folder
    if(!row.Path.includes('All Files/HomeDir Staff/CareyG/')){
      //Check for invalid characters
      if(row.ItemName.includes('?') ||
      row.ItemName.includes('*') ||
      row.ItemName.includes(':') ||
      row.ItemName.includes('|') ||
      row.ItemName.includes('\"') ||
      row.ItemName.includes('<') ||
      row.ItemName.includes('>') ||
      row.ItemName.includes('/') ||
      row.ItemName.includes('\\') ||
      row.ItemName.includes('_vti_')
      ){
        logger.warn('Invalid character: ' + row.ItemID + ' ' + row.ItemType + ' ' + row.ItemName);

        //Replace with permitted character
        var newName = row.ItemName.replace('?', ' ');
        newName = newName.replace('*', ' ');
        newName = newName.replace(':', '-');
        newName = newName.replace('|', '-');
        newName = newName.replace('\"', '\'');
        newName = newName.replace('<', '');
        newName = newName.replace('>', '');
        newName = newName.replace('/', '');
        newName = newName.replace('\\', '');
        newName = newName.replace('_vti_', 'vti-removed-in-migration');

        if(row.OwnerLogin.includes('boxadmin@krb.nsw.edu.au')){
          if(!row.ItemName.includes('#NAME?')){
            if(row.ItemType.includes('File')){
              updateBoxFile(row.ItemID, newName);
            } else if (row.ItemType.includes('Folder')){
              updateBoxFolder(row.ItemID, newName);
            } else {
              logger.error('Unknown object type: ' + row.ItemID + ' ' + row.ItemName);
            }
          } else {
            logger.warn('Ignoring #NAME? from xlsx to csv conversion: '+ row.ItemID);
          }
        } else {
          logger.error('Object not owned by boxadmin and needs a new name: ' + row.ItemID + ' ' + row.OwnerLogin);
        }

      }
    } 

    //You can’t have folder names in SharePoint Online that begins with a tilde (~).
    if(row.ItemName.startsWith('~') & row.ItemType.includes('Folder')){
      logger.warn('Invalid folder starting with ~: ' + row.ItemID + ' ' + row.ItemName);
      var newName = row.ItemName.replace('~', 'migrated-');
      //replace tilde with 'migrated-...'
      if(row.OwnerLogin.includes('boxadmin@krb.nsw.edu.au')){
        updateBoxFolder(row.ItemID, newName);
      } else {
        logger.error('Object not owned by boxadmin and needs a new name: ' + row.ItemID + ' ' + row.OwnerLogin);
      }
    }

    // The entire path, including the file name, must contain fewer than 400 characters
    if(row.Path.length > 400){
      logger.error('Path name is longer than 400 chars: ' + row.ItemID + ' ' + row.Path );
      //what do we do?
    }

    //15GB File upload size limit
    if(row.Size.includes('GB') & row.ItemType.includes('File')){
      var NUMERIC_REGEXP = /[-]{0,1}[\d]*[\.]{0,1}[\d]+/g;
      var gigabytes = parseFloat(Number(row.Size.match(NUMERIC_REGEXP)));
      if(gigabytes > 15.0) {
        logger.error('File larger than 15GB: ' + row.ItemID + ' ' + row.OwnerLogin + row.ItemName + row.Size);
        //what do we do?
      }
    }
  })
  .on('end', () => {
    logger.info('CSV file successfully processed');
  });


// Call Box SDK API
function updateBoxFile(boxID, newFileName){
  client.files.update(boxID, {name : newFileName})
	.then(updatedFile => {
		logger.info('client.files.update: ' + boxID + ' ' + newFileName);
	})
  .catch(err => logger.error('Box file API POST error: ' + boxID + ' ' + err));

  // client.files.get(boxID)
  //     .then(file => {
  //       //logger.info('Box file to rename: ' + boxID + ' ' + file.name);
  //       logger.info('client.files.update: ' + boxID + ' ' + newFileName);
  //     })
  //     .catch(err => logger.error('Box file API GET error: ' + boxID + ' ' + err));
}

function updateBoxFolder(boxID, newFolderName){
  client.folders.update(boxID, {name : newFolderName})
	.then(updatedFolder => {
		logger.info('client.folders.update: ' + boxID + ' ' + newFolderName);
	})
  .catch(err => logger.error('Box folder API POST error: ' + boxID + ' ' + err));
  // client.folders.get(boxID)
  // 	.then(folder => {
  //     //logger.info('Box folder to rename: ' +  folder.name);
  //     logger.info('client.folders.update: ' + boxID + ' ' + newFolderName);
  // 	})
  //   .catch(err => logger.error('Box folder API GET error: ' + boxID + ' ' + err));
}
