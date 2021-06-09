/*Enable 'Use Lambda Proxy integration'*/

'use strict';
const AWS = require('aws-sdk');

const ddb = new AWS.DynamoDB.DocumentClient({ region: 'us-east-2' });

exports.handler = function(event, context, callback) {
 if (event.queryStringParameters && event.queryStringParameters.licensekey && event.queryStringParameters.greeter !== "") {

  var key = event.queryStringParameters.licensekey;

  var params = {
   TableName: 'licenses',
   Key: {
    'licensekey': key
   }
  }
  ddb.get(params, function(err, data) {
   if (err) {
    callback(err, null);
   }
   else {

    var response = {
     'body': JSON.stringify(data),
     "headers": {
      "Content-Type": "application/json"
     },
     "statusCode": 200
    };
    callback(null, response);
   }
  })
 }
 else {
  const user = {
   Item: {
    res:'missing parameters'
   }
  };
  var response = {
   'body': JSON.stringify(user),
   "headers": {
    "Content-Type": "application/json"
   },
   "statusCode": 200
  };
  callback(null, response);
 }

}
