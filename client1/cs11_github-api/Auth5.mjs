
     import * as MSAuthLib from '@azure/msal-node';
     import { inspect    } from 'util'

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

//          aAPIpath =  'https://graph.microsoft.com/v1.0/users/{userPrincipalName}/onenote/{notebooks | sections | sectionGroups | pages}' 
//          aAPIpath =  'https://graph.microsoft.com/v1.0/groups/{id}/onenote/notebooks'
//          aAPIpath =  'https://graph.microsoft.com/v1.0/sites/{id}/onenote/notebooks'
//          aAPIpath =  'https://graph.microsoft.com/beta/teams/{group-id-for-teams}/channels/{channel-id}/messages'

       var  bToken   =  (String(mScopes).length > 500)  // == 3184 

            aAPIpath =   aAPIpath.replace( /^\//, '' )
            aAPIpath =   aAPIpath.replace( /{id}/, aOwnerId ? aOwnerId.trim() : '' )
//          aAPIpath =  `https://graph.microsoft.com/v1.0/${ aAPIpath }`
//          aAPIpath =  `https://graph.microsoft.com/beta/${ aAPIpath }`
            aAPIpath =  `https://graph.microsoft.com/{vn}/${ aAPIpath }`
            aAPIpath =   aAPIpath.replace( /{vn}/, aVersion ? aVersion : 'v1.0' )

        if (bToken == false) {
            mScopes  =           mScopes              ? mScopes              : ''
            mScopes  =   typeof( mScopes) == 'string' ? mScopes.split( ',' ) : mScopes
            mScopes  =           mScopes[0] == ''     ? [ '.default' ]       : mScopes 
            mScopes  =   mScopes.map( aScope => `https://graph.microsoft.com/${ aScope }` )
            }
//     var  aDate    =  fmtDate( 22 )
       var  aMsg     =  `\n${ fmtDate( 16 ).substring(  2 ) } `.padEnd( 107, '-' )
            aMsg    +=  `\n${ fmtDate( 22 ).substring( 14 ) } Submitting MSGraph API -\n` 
      try {  
        if (bToken) {    var aLen = `${ aToken.length }` 
            aMsg    +=`  Token:  ${ aToken.substring( 0, 90 - (aLen.length + 6) ) }... (${ aToken.length })` 
       var  pTokens  = { accessToken: mScopes }
        } else {
            aMsg    +=`  Scopes: ${ mScopes.join( '\n'.padEnd(8) ) }` 
       var  pTokens  =   await  authClient.acquireTokenByClientCredential( { scopes: mScopes } )
            }
        if (aAPIpath) {
       var  pHeaders = {'Authorization': `Bearer ${ pTokens.accessToken }` }

       var  pResult  =   await  fetch( aAPIpath, { headers: pHeaders } )
       var  pJSON    =   await  pResult.json( );
        } else {
            pJSON    = { }
            }
        if (pJSON.value) {     
            pJSON    = { token: pTokens.accessToken, '@odata.context': pJSON['@odata.context'], value: pJSON.value }
       } else { 
            pJSON    = { error: { code: pJSON.error.code, message: pJSON.error.message }
//                     , token:   pTokens ? pTokens.accessToken.substring( 0,100) + `... (${pTokens.accessToken})`: '', value: [] } } 
                       , token:   pTokens ? fmtToken( pTokens ) : '', value: [] }
            }
   } catch( pErr ) {
            pJSON    = { error: { code: pErr.errorCode, message: pErr.errorMessage }
                       , token:   pTokens ? fmtToken( pTokens ) : '', value: [] }
            }
            console.log( aMsg + `\n  URL:    ${aAPIpath}` )
     return pJSON 

  function  fmtToken( pTokens ) {
       var  aToken   =   pTokens ? pTokens.accessToken : '', aLen = `${aToken.length}`
    return `${ aToken.substring( 0, 90 - (aLen.length + 6) ) }... (${ aToken.length })` 
            }    
            }
// ------------------------------------------------------------------

     async  function getAPI(  aPath, aScopes, shoResult_, aId, aIdName, aVersion ) {
        if (typeof(aId) == 'function') { var x = aId; aId = shoResult_; shoResult = x }

       var  pResult  =  await queryGraphApi(  aPath, aScopes, aId, aVersion )

        if (pResult.error) {  return pResult }
      
            pResult.data   =  shoResult_( pResult.value, "asData" )
             
            console.log( `${  fmtDate( 22 ).substring( 14 ) }  ${ aIdName ? `${ aIdName.trim() } ` : "" }Complete. `.padEnd( 87, "-" ) + "\n" )

     return pResult
            }  // eof getResult
// -------  -----------------------------------------------------------------------------------------------------------------------

     async  function shoAPI(  aPath, aScopes, shoResult_, aId, aIdName, aVersion ) {
        if (typeof(aId) == 'function') { var x = aId; aId = shoResult_; shoResult_ = x }

       var  pResult  =  await queryGraphApi(  aPath, aScopes, aId, aVersion ); // if (aIdName) {
            console.log(  aIdName ?  "  id:".padEnd(10) + aIdName.trim() + "\n" : "" ) // }

        if (pResult.error) {  var aCR = "\n".padEnd(11)
            console.log(     `Error: ** ${  pResult.error.code }${aCR}${ pResult.error.message.replace( /\n/g, aCR ) }` )
        } else {
//      if (shoResult_) {
//          pResult.data   =  shoResult_( pResult.value,      'asData' ) }
            pResult.data   =  shoResult( pResult, shoResult_, 'asData' ) }
            console.log(      shoResult( pResult, shoResult_           ) )  // if shoResult1, then asData = 'byRows'  
//          }
            console.log( `${  fmtDate( 22 ).substring( 14 ) }  ${ aIdName ? `${ aIdName.trim() } ` : "" }Complete. `.padEnd( 87, "-" ) + "\n" )

    return  pResult

//    ----  ---------------------------------------------

  function  shoResult( pJSON, shoJSON, asData ) {
       var  aFncName=  String( shoJSON ).match( /function +(.+?)\(/ )[1] 
            asData  =  asData ? asData : ( aFncName.match(/1$/) ? 'asCSV' : '') 
        if (pJSON [   '@odata.context' ]) {                                   // API URL
//      if (asData == 'asData') { return pJSON.value }                                  // original data, not fmtData1 
        if (shoJSON) { 
        if (asData == 'asData') { return              shoJSON( pJSON.value,'asData') }  // fmtDataX asdata  
    return  `Result:   ${aFncName} ( pJSON.value )` + shoJSON( pJSON.value, asData )
         }  }
        if (asData == 'asData') { return pJSON }
//  return  'Result:   pJSON.value\n' + inspect( pJSON.value, { depth: 9 } )  // If not pJSON [ '@odata.context' ], does pJSON.value exist
    return  'Result:   pJSON.value\n' + inspect( pJSON,       { depth: 9 } )
            }  // eof shoResult
//    ----  ---------------------------------------------
         };  // eof shoAPI
// -------------------------------------------------------------------------

function  fmtData( mData, asData, aName, bCSV ) {
       var  aData =  asData ? asData : 'asCSV';                  // was byRows
        if (aData =='asCSV') { aData = 'byRows'; bCSV = true } 
       var  bQQ   = (bCSV == "qq"); bCSV = bQQ ? false : bCSV

        if (aData =='asData' ) { return mData }

        if (aData =='byFlds' ) {
       var  mRecs =  mData.map( pFlds => Object.keys( pFlds ).map( aFld => `${ take( 20, aFld ) }: ${ pFlds[ aFld ] }` ).join( "\n  " ) )
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
       var  aRecs =  aLine + mRecs.join( aLine ) + aLine.replace( /[\n ]+$/, "" )
            }
        if (aData =='byRows' ) {
       var  mRecs =  mData.map( pFlds => Object.keys( pFlds ).map( aFld => qq( pFlds[ aFld ], bCSV ) ).join( bCSV ? "," : " " ) )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            }
            aRecs = (aName ? `\n${aName}:` : '') + aRecs
    return  aRecs

  function  qq( a, bCSV ) { a = String( a )
    return (/ /.test(a) && bCSV) ? `"${ a.replace( /"/g, '\\"' )}"` : ( a ? (bQQ ? a.replace( /'/g, "\\'" ) : a ) : ( bCSV ? '' : "''" ))
            }
         }; // eof fmtData
// -------------------------------------------------------------------------

  function  fmtDate( n, a ) {
        if (typeof(n) != 'number') { var x = a; a = n; n = x }; n = n ? n : 19
            a = ( typeof(a) != 'undefined' ) ? ( String(a).match( /Time/) ? a.toISOString() : a ) : new Date().toISOString();   // no typeof(a) == 'date'
    return  a.replace( /[TZ]/g, ' ' ).substring( (n == 8 ? 2 : 0), (n == 8 ? 10 : n ) )
            }
//    ----  ---------------------------------------------
// -------------------------------------------------------------------------


    export { queryGraphApi as queryGraphApi_v105 };
    export { getAPI, shoAPI, fmtData };


