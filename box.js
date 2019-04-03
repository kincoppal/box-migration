// Testing of Box SDK
const config = require('./config');
const BoxSDK = require('box-node-sdk');

// Initialize the SDK with your app credentials
var sdk = new BoxSDK({
  clientID: config.boxAuth.clientID,
  clientSecret: config.boxAuth.clientSecret
});

// Create a basic API client, which does not automatically refresh the access token
var client = sdk.getBasicClient(config.boxAuth.developerToken);

// Get your own user object from the Box API
// All client methods return a promise that resolves to the results of the API call,
// or rejects when an error occurs
client.users.get(client.CURRENT_USER_ID)
	.then(user => console.log('Hello', user.name, '!'))
	.catch(err => console.log('Got an error!', err));

// client.folders.get('47974880975')
//   .then(folder => {
//     console.log('Folder Name:', folder.name);
//     console.log('Size:', folder.size);
//   })
//   .catch(err => console.log('API error:', err.statusCode));

client.files.get('252151559586')
	.then(file => {
		// ...
    console.log('File name:', file.name);
	})
  .catch(err => console.log('API error:', err.statusCode));

client.folders.get('42667180820')
	.then(folder => {
		// ...
    console.log('Folder name:', folder.name);
	})
  .catch(err => console.log('API error:', err));

// client.files.update('433000352883', {name : 'History 2 - Stage 3.pptx'})
// 	.then(updatedFile => {
// 		// ...
//     console.log('File name has been updated')
//     client.files.get('433000352883')
//     	.then(file => {
//     		// ...
//         console.log('File name:', file.name);
//     	})
//       .catch(err => console.log('API error:', err.statusCode));;
// 	})
//   .catch(err => console.log('API error:', err.statusCode));;
