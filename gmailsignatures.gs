

function getDomainUsersList() {
  
  var users = [];
  var options = {
    domain: 'domain.co.uk',     // Google Apps domain name
    customer: 'my_customer',        //Applies to most people. Partners would put their customers id
    maxResults: 500,
    projection: 'custom',      
    query: "orgUnitPath='/Current Staff'",  //Restricts this to my Current Staff OU
    viewType: 'admin_view',
    orderBy: 'email'          // Sort results by users
  }
  
  do {
    var response = AdminDirectory.Users.list(options);
    response.users.forEach(function(user) {
      users.push([user.name.fullName, user.primaryEmail, user.organizations]);  //gets name,email and organization details for each staff member
    });
    
    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);
  
  // Insert data in a spreadsheet
  var ss = SpreadsheetApp.openById('');
  var sheet = ss.getSheetByName('Users') || ss.insertSheet('Users', 1);
  sheet.clear();
  //Add some column headings
  sheet.getRange('A1').setValue('JobTitle')
  sheet.getRange('B1').setValue('Name')
  sheet.getRange('C1').setValue('Email')
  sheet.getRange('D1').setValue('Organizations')
  sheet.getRange(2,2,users.length, users[0].length).setValues(users);
  sheet.getRange('A2').activate();
  //insert a formula that pulls out job title from the organisation column
  sheet.getCurrentCell().setFormula('=IFERROR(mid(D2,find("title=",D2)+6,find(",",D2,find("title=",D2))-find("title=",D2)-6))');
  sheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet.getRange('E1').setValue('DisplayName')
  sheet.getRange('E2').activate();
  //insert a formula that pulls out the department field.  In the department field I am syncing displayname from active directory.
  sheet.getCurrentCell().setFormula('=IFERROR(mid(D2,find("department=",D2)+11,find(",",D2,find("department=",D2))-find("department=",D2)-11))');
  sheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
   
  
  // I was having trouble with some users where in the D column cell their department was the last thing in the array.
  //  So add an extra comma to the end of the sting to make the above formula work.
  
  var values = sheet.getDataRange().getValues();  //find the range of the used cells
  for (var j = 1; j <= values.length; j++) {  //iterate through all used rows
   //Get the cell value in row J column 4.
  var cellvalue = sheet.getRange(j, 4).getValue()
    //For any values in organisation that don't have a comma after the Department.
    var tempstring = cellvalue.toString().replace('}',',}');
// Updating text on a particular cell
sheet.getRange(j,4).setValue(tempstring)
  }
  
 //Hide the big organisation column - not really necessary
  sheet.getRange('D:D').activate();
  sheet.hideColumns(sheet.getActiveRange().getColumn(), sheet.getActiveRange().getNumColumns());
  
  //Resize the columns to make it look good
  sheet.autoResizeColumns(1, sheet.getMaxColumns());

  
  //Now I have all the info I need to make my signature.
}


//credentials is the contents of a json private key file for a service account with domain wide authentication
//client_id is added to the oauth scopes in Google Admin, Security, Advanced, Managed oauth scopes.
//grant access to 'https://www.googleapis.com/auth/gmail.settings.basic'

var credentials = {
  'type': 'service_account',
  'project_id': 'project-id-',
  'private_key_id': '',
  'private_key': '',
  'client_email': 'gmailsignatures@project-id-.iam.gserviceaccount.com',
  'client_id': '',
  'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
  'token_uri': 'https://accounts.google.com/o/oauth2/token',
  'auth_provider_x509_cert_url': 'https://www.googleapis.com/oauth2/v1/certs',
  'client_x509_cert_url': 'https://www.googleapis.com/robot/v1/metadata/x509/gmailsignatures%40project-id4.iam.gserviceaccount.com'
}


  


function setSignatureLoop() {

  //this loops through first 175 records in sheet and assigns the signature
  //175 to avoid going over execution time
  
  var ss = SpreadsheetApp.openById('sheetid');
  var sheet = ss.getSheetByName('Users');
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < 175; i++) {
    var email = values[i][2];
    var displayname = values[i][4];
    var jobtitle = values [i][0];
   // if the job title is Shared Mailbox a slightly different signature goes in
    if (jobtitle == 'Shared Mailbox') {
      //signature in HTML with displayname inserted
       var signature = "<div dir='ltr'><font style='color:rgb(25,35,93)' size='2'><b>" + displayname + "</b></font><div><font style='color:rgb(25,35,93)' size='2'>School name</font></div><div style='color:rgb(25,35,93)'><font size='2'>Address</div><div style='color:rgb(25,35,93)'><font size='2'>tel no| <a href='http://www.website.com' style='color:rgb(25,35,93)' target='_blank'>www.website.com</a></font></div><div><font size='1'> </font></div><div><img src='https://drive.google.com/a/LINKTOLOGO;export=download' width='300' height='75'><br></div></div>"
      var setSig = setSignature(email, signature);
    }
    else {
      var signature = "<div dir='ltr'><font style='color:rgb(25,35,93)' size='2'><b>" + displayname + " | " + jobtitle + "</b></font><div><font style='color:rgb(25,35,93)' size='2'>School Name</font></div><div style='color:rgb(25,35,93)'><font size='2'>Address</div><div style='color:rgb(25,35,93)'><font size='2'>tel no | <a href='http://www.website.com' style='color:rgb(25,35,93)' target='_blank'>www.website.com</a></font></div><div><font size='1'> </font></div><div><img src='https://drive.google.com/a/LINKTOLOGO;export=download' width='300' height='75'><br></div></div>"
      var setSig = setSignature(email, signature);
    }
  }
}
//doing the rest of the loop.  You may need more of these if your current staff list is bigger.
function setSignatureLoop175toEnd() {

  
  var ss = SpreadsheetApp.openById('sheetid');
  var sheet = ss.getSheetByName('Users');
  var values = sheet.getDataRange().getValues();
  
  for (var i = 175; i < values.length; i++) {
    var email = values[i][2];
    var displayname = values[i][4];
    var jobtitle = values [i][0];
    if (jobtitle == 'Shared Mailbox') {
      var signature = "<div dir='ltr'><font style='color:rgb(25,35,93)' size='2'><b>" + displayname + "</b></font><div><font style='color:rgb(25,35,93)' size='2'>School name</font></div><div style='color:rgb(25,35,93)'><font size='2'>Address</div><div style='color:rgb(25,35,93)'><font size='2'>tel no| <a href='http://www.website.com' style='color:rgb(25,35,93)' target='_blank'>www.website.com</a></font></div><div><font size='1'> </font></div><div><img src='https://drive.google.com/a/LINKTOLOGO;export=download' width='300' height='75'><br></div></div>"
      var setSig = setSignature(email, signature);
    }
    else {
      var signature = "<div dir='ltr'><font style='color:rgb(25,35,93)' size='2'><b>" + displayname + " | " + jobtitle + "</b></font><div><font style='color:rgb(25,35,93)' size='2'>School Name</font></div><div style='color:rgb(25,35,93)'><font size='2'>Address</div><div style='color:rgb(25,35,93)'><font size='2'>tel no | <a href='http://www.website.com' style='color:rgb(25,35,93)' target='_blank'>www.website.com</a></font></div><div><font size='1'> </font></div><div><img src='https://drive.google.com/a/LINKTOLOGO;export=download' width='300' height='75'><br></div></div>"
      var setSig = setSignature(email, signature);
    }
  }
}



function setSignature(email, signature) {


  var signatureSetSuccessfully = false;

  var service = getDomainWideDelegationService('Gmail: ', 'https://www.googleapis.com/auth/gmail.settings.basic', email);

  if (!service.hasAccess()) {

    Logger.log('failed to authenticate as user ' + email);

    Logger.log(service.getLastError());

    signatureSetSuccessfully = service.getLastError();

    return signatureSetSuccessfully;

  } else Logger.log('successfully authenticated as user ' + email);

  var username = email.split("@")[0];


  var resource = { signature: signature };

  var requestBody                = {};
  requestBody.headers            = {'Authorization': 'Bearer ' + service.getAccessToken()};
  requestBody.contentType        = "application/json";
  requestBody.method             = "PUT";
  requestBody.payload            = JSON.stringify(resource);
  requestBody.muteHttpExceptions = false;

  var emailForUrl = encodeURIComponent(email);

  var url = 'https://www.googleapis.com/gmail/v1/users/me/settings/sendAs/' + emailForUrl;

  
  var maxSetSignatureAttempts     = 2;
  var currentSetSignatureAttempts = 0;

  
  do {

try {

      currentSetSignatureAttempts++;
   //   Logger.log('currentSetSignatureAttempts: ' + currentSetSignatureAttempts);
      var setSignatureResponse = UrlFetchApp.fetch(url, requestBody);
   //   Logger.log('setSignatureResponse on successful attempt:' + setSignatureResponse);
      signatureSetSuccessfully = true;
      break;

   } catch(e) {

   // Logger.log('set signature failed attempt, waiting 3 seconds and re-trying');

      Utilities.sleep(3000);

  }

    if (currentSetSignatureAttempts >= maxSetSignatureAttempts) {

      Logger.log('exceeded ' + maxSetSignatureAttempts + ' set signature attempts, deleting user and ending script');


   }
  } while (!signatureSetSuccessfully);

  return signatureSetSuccessfully;

}


function getDomainWideDelegationService(serviceName, scope, email) {

  //Logger.log('starting getDomainWideDelegationService for email: ' + email);

  return OAuth2.createService(serviceName + email)
      // Set the endpoint URL.
      .setTokenUrl(credentials.token_uri)

      // Set the private key and issuer.
      .setPrivateKey(credentials.private_key)
      .setIssuer(credentials.client_email)

      // Set the name of the user to impersonate. This will only work for
      // Google Apps for Work/EDU accounts whose admin has setup domain-wide
      // delegation:
      // https://developers.google.com/identity/protocols/OAuth2ServiceAccount#delegatingauthority
      .setSubject(email)

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getScriptProperties())

      // Set the scope. This must match one of the scopes configured during the
      // setup of domain-wide delegation.
      .setScope(scope);

}
