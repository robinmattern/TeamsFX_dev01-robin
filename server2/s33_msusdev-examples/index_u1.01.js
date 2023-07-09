async function run() {

        console.log( "running..." );

    const config = {
        auth: {
            clientId   : '0d2192da-b16a-4f6b-8378-8cd1c08a799d',
            authority  : 'https://login.microsoftonline.com/b5d500c1-e385-42eb-ab34-db10f116a192/',
            redirectUri: 'http://localhost:8080'
            }
        };
    var client = new msal.PublicClientApplication( config );
    
    var loginRequest = {
        scopes: [ 'user.read' ]
        };

    let loginResponse = await client.loginPopup( loginRequest );
        console.log('Login Response', loginResponse);

    var tokenRequest = {
        scopes: [ 'user.read' ],
        account: loginResponse.account
        };

    let tokenResponse = await client.acquireTokenSilent( tokenRequest );
        console.log('Token Response', tokenResponse);

    let payload = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
            'Authorization': 'Bearer ' + tokenResponse.accessToken
            }
        });

    let json = await payload.json();
        console.log( 'Graph Response', json );
}