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
    var curdate = new Date();

var validuntill = new Date(data.Item['validuntil']);
 var isvalid = curdate.getTime() <= validuntill.getTime(); 
    
    
    const res = {
     valid : data.Item['validuntil'],
     date : curdate,
     durum: isvalid,
     envanter: data.Item['envanter'],
     risk: data.Item['risk'],
     denetim: data.Item['denetim']
    };
    var response = {
     'body': JSON.stringify(res),
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
    res: '404'
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

