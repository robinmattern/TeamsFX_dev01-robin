
     import * as MSAuthLib from '@azure/msal-node';

//   import fetch from 'node-fetch';
//   https://jwt.ms/   // View token contents 
//   https://dev.to/davelosert/practical-guide-to-use-the-microsoft-graph-api-4ahn
//   https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/0d2192da-b16a-4f6b-8378-8cd1c08a799d/isMSAApp~/false

            process.env.TenantId     = 'b5d500c1-e385-42eb-ab34-db10f116a192';     // SicommNet, Inc 
            process.env.ClientId     = '0d2192da-b16a-4f6b-8378-8cd1c08a799d';     // FRApps-s38_teams-msgraph-api_dev03-robin
            process.env.ClientSecret = 'fIK8Q~Ef71EHU4DV7JAH4UqATC_7LEIb0wjIbdy5';

       var  pClientConfig             =
             {  auth:
                 {  clientId         :  process.env.ClientId
                 ,  clientSecret     :  process.env.ClientSecret
                 ,  authority        : `https://login.microsoftonline.com/${ process.env.TenantId  }`
                    }
                };

       var  authClient =   new MSAuthLib.ConfidentialClientApplication( pClientConfig );

       var  aToken     =  'eyJ0eXAiOiJKV1QiLCJub25jZSI6InZSd1c2ZWhCTms2RVhXYkFEWWVNb2hILUhpUTd6MzdLbkdZa09yd0J2ek0iLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIvIiwiaWF0IjoxNjg4MDY0OTM1LCJuYmYiOjE2ODgwNjQ5MzUsImV4cCI6MTY4ODE1MTYzNSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhUQUFBQXl1U3VlQkxGUy9QVWtERGVYY0x4dTNjWXVCL2pFMVF1THNMcU1naEo0a2k2dWV3d3RrSzlVYTVvbDVPQmNMNDdQQzRRd3dudVhySmdPUFdINExZd3RtbnhKd2lzRlVpamlNZDZMU1YyS1dBPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiTWF0dGVybiIsImdpdmVuX25hbWUiOiJSb2JpbiIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjczLjEzMy44NC4xMDMiLCJuYW1lIjoiUm9iaW4gTWF0dGVybiIsIm9pZCI6ImIwMGY3ZDIxLWJhOTYtNDFkOC05YjIxLTg4NWIwNjc4ZDc0OSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzQkZGRDhDM0UyMEM4IiwicmgiOiIwLkFUY0F3UURWdFlYajYwS3JOTnNROFJhaGtnTUFBQUFBQUFBQXdBQUFBQUFBQUFBM0FPMC4iLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlByaXZpbGVnZWRPcGVyYXRpb25zLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BbGwgR3JvdXAuUmVhZFdyaXRlLkFsbCBJZGVudGl0eVJpc2tFdmVudC5SZWFkLkFsbCBNYWlsLlJlYWRXcml0ZSBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLlJlYWQgTm90ZXMuUmVhZC5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgUmVwb3J0cy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWQgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6Ijd6blpjdzFDRjgxa0JzZE9IN3Y5U0owY0tZZi1LdFdaWm5HZmthMzYzTVkiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiJiNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIiLCJ1bmlxdWVfbmFtZSI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInVwbiI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInV0aSI6IkMwMmphN3A1VlVtTTMtanZUcWxkQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfY2MiOlsiQ1AxIl0sInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6Ik1ENkFTb2J3ZXhTX0o4TmJIN0JJd29LdHpMTGtudzJxdUsxYzNEWlp4XzQifSwieG1zX3RjZHQiOjE0MTI4MDU0MDV9.LHQeYWxUaPn0K9U-87tsMYF9rweFanWAUBN0FSoerUIjzcBOk_WWVEFkT4iavKUEkTR0PqHNZ_6zFGguX5v_M4SXlwf5PDkDsiCxmiVfDyu_m0OcwGgyF4Uisu455kr5fZuGfXE9sDChdpE9_H3elNwrh3TUHn5IfXt5s7nJnsoWD8v2WNRwZZHC5WjTp7StwBc_IWds20jzb4YReIWqtT7GCZ5hllCQv94tedF9dSDKJLgj8m76ym1LehehnvfDNTRIxAhSTMDsj43y33bdBrX6SC9m8ZI0pAmr5QlOINHPFpcdkqozn7lRF0viZOUOwfBeXGY7fMOSEmvIOkZJ1A'

//          queryGraphApi( aToken ) 

// -----------------------------------------------------------------

     async  function queryGraphApi( aAPIpath, mScopes, aOwnerId, aVersion ) {

//          aAPIpath  = 'https://graph.microsoft.com/v1.0/users/{userPrincipalName}/onenote/{notebooks | sections | sectionGroups | pages}' 
//          aAPIpath  = 'https://graph.microsoft.com/v1.0/groups/{id}/onenote/notebooks'
//          aAPIpath  = 'https://graph.microsoft.com/v1.0/sites/{id}/onenote/notebooks'
//          aAPIpath  = 'https://graph.microsoft.com/beta/teams/{group-id-for-teams}/channels/{channel-id}/messages'

       var  bToken    = (String(mScopes).length > 500)  // == 3184 

            aAPIpath  =  aAPIpath.replace( /^\//, '' )
            aAPIpath  =  aAPIpath.replace( /{id}/, aOwnerId ? aOwnerId : '' )
//          aAPIpath  = `https://graph.microsoft.com/v1.0/${ aAPIpath }`
//          aAPIpath  = `https://graph.microsoft.com/beta/${ aAPIpath }`
            aAPIpath  = `https://graph.microsoft.com/{vn}/${ aAPIpath }`
            aAPIpath  =  aAPIpath.replace( /{vn}/, aVersion ? aVersion : 'v1.0' )

        if (bToken == false) {
            mScopes   =          mScopes              ? mScopes              : ''
            mScopes   =  typeof( mScopes) == 'string' ? mScopes.split( ',' ) : mScopes
            mScopes   =          mScopes[0] == ''     ? [ '.default' ]       : mScopes 
            mScopes   =  mScopes.map( aScope => `https://graph.microsoft.com/${ aScope }` )
            }
      try {  

        var aMsg      = '\nSubmitting MSGraph API -\n' 
        if (bToken) {
       var  pTokens   ={ accessToken: mScopes }
            aMsg     +=` Token: ${   aToken.substring( 0, 80 ) }... (${aToken.length})\n URL:    ${aAPIpath}` 
        } else {
       var  pTokens   =  await  authClient.acquireTokenByClientCredential( { scopes: mScopes } )
            aMsg     +=` Scopes: ${ mScopes.join('\n       ' ) }\n URL:    ${aAPIpath}` 
            }

       var  pHeaders  ={'Authorization': `Bearer ${ pTokens.accessToken }`}

       var  pResult   =  await  fetch( aAPIpath, { headers: pHeaders } )
       var  pJSON     =  await  pResult.json( );




            console.log( aMsg )

   } catch( pErr ) {
            pJSON     ={ error: { code: pErr.errorCode, message: pErr.errorMessage } }

            console.log( aMsg )
            }
     return pJSON 
            }
// ------------------------------------------------------------------

    export { queryGraphApi };


