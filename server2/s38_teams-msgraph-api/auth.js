//MSAL configuration

    const  msalConfig =
            {  auth:
                {  clientId:    '0d2192da-b16a-4f6b-8378-8cd1c08a799d'
                ,  authority:   'https://login.microsoftonline.com/b5d500c1-e385-42eb-ab34-db10f116a192'  // comment out if you use a multi-tenant AAD app
                ,  redirectUri: 'http://localhost:8080'
                   }
               };

     const  msalRequest     = { scopes: [ ] };

// ------------------------------------------------------------------------------

  function  ensureScope( aScopeRequest ) {     // called by Microsoft Signin Dialog

       var  bOK             =  msalRequest.scopes.some( ( aScopeSaved ) => {
                        return aScopeSaved.toLowerCase() === aScopeRequest.toLowerCase( )
                               } )
       if (!bOK) {             msalRequest.scopes.push(   aScopeRequest ) }
            msalRequest.scopes.map( aScopeSaved  => {
            console.log(          `  scopeRequest:      ${ aScopeSaved }` ) } )
            }
// ------------------------------------------------------------------------------

     const  msalClient      =  new msal.PublicClientApplication( msalConfig ); // Initialize MSAL client

// ------------------------------------------------------------------------------

     async  function signIn( ) { // Log the user in

     const  authResult      =  await  msalClient.loginPopup( msalRequest );
            sessionStorage.setItem(  'msalAccount', authResult.account.username );
            }
// ------------------------------------------------------------------------------

     async  function getToken( ) { //Get token from Graph

       let  account         =         sessionStorage.getItem( 'msalAccount' );
            console.log(           `  account:           ${     account  }`  )

       if (!account) {
            throw new Error( 'User info cleared from session. Please sign out and sign in again.'); }
      try {
     const  silentRequest   =     // First, attempt to get the token silently
             {  scopes      :         msalRequest.scopes
             ,  account     :         msalClient.getAccountByUsername( account )
                };

     const  silentResult    =  await  msalClient.acquireTokenSilent(   silentRequest );
//          console.log(           `  silentResult:      ${ JSON.stringify(   silentResult  ) }` )
            console.log(           `  silentResultToken: ${ silentResult.accessToken.substring( 0, 50 ) }...` )
            console.log(           `  scopesResult:      ${ silentResult.scopes      }` )

    return  silentResult.accessToken;

       } catch ( silentError ) {                    // If silent requests fails with InteractionRequiredAuthError,

        if (silentError instanceof msal.InteractionRequiredAuthError) { // attempt to get the token interactively

     const  interactiveResult= await  msalClient.acquireTokenPopup(    msalRequest   );
//          console.log(           `  interactiveResult: ${ JSON.stringify(  interactiveResult ) }` )
            console.log(           `  interaResultToken: ${ interactiveResult.substring( 0, 50 ) }...` )
            console.log(           `  scopesResult:      ${ interactiveResult.scopes      }` )

    return  interactiveResult.accessToken;

        } else { throw silentError; }
      }
    }
// ------------------------------------------------------------------------------

/*
    `npm start`

        > mslearn-retrieve-m365-data-with-msgraph-quickstart@1.0.0 start
        > npx http-server -c-1 -a localhost -o

        Starting up http-server, serving ./

        http-server version: 14.1.1

        http-server settings:
        CORS                    : disabled
        Cache                   : -1 seconds
        Connection Timeout      : 120 seconds
        Directory Listings      : visible
        AutoIndex               : visible
        Serve GZIP Files        : false
        Serve Brotli Files      : false
        Default File Extension  : none

        Available on            : http://localhost:8080

        Hit CTRL-C to stop the server
        Open: http://localhost:8080

        (node:31444) [DEP0066] DeprecationWarning: OutgoingMessage.prototype._headers is deprecated
        (Use `node --trace-deprecation ...` to show where the warning was created)

        [2023-06-28T18:41:56.342Z]  "GET /"                                         "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:41:56.658Z]  "GET /images/ms-symbollockup_signin_light.png"  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:41:56.757Z]  "GET /auth.js"                                  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:41:56.767Z]  "GET /graph.js"                                 "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:41:56.864Z]  "GET /ui.js"                                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"

        [2023-06-28T18:42:04.009Z]  "GET /"                                         "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:04.066Z]  "GET /images/ms-symbollockup_signin_light.png"  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:04.072Z]  "GET /auth.js"                                  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:04.077Z]  "GET /graph.js"                                 "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:04.080Z]  "GET /ui.js"                                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"

        [2023-06-28T18:42:11.106Z]  "GET /"                                         "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:11.167Z]  "GET /images/ms-symbollockup_signin_light.png"  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:11.171Z]  "GET /auth.js"                                  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:11.178Z]  "GET /graph.js"                                 "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
        [2023-06-28T18:42:11.186Z]  "GET /ui.js"                                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.58"
*/