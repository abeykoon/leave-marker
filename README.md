# Leave-marker
Ballerina program to mark leaves of wso2 team members

This program will scan through emails, find emails send notifying the vacations and ultimately mark them in a Google Shhet configured. 
Following the the configurables to set. 

Obtain Gmail and GSheet tokens following the guide at [Google](https://developers.google.com/identity/protocols/oauth2)

```
[gmailOAuthConfig]
clientId=""
clientSecret=""
refreshToken=""

[gsheetOAuthConfig]
clientId=""
clientSecret=""
refreshToken=""

spreadSheetId=""
workSheetName=""
```
