
     import * as msal from '@azure/msal-node';

//   import fetch from 'node-fetch';
//   https://dev.to/davelosert/practical-guide-to-use-the-microsoft-graph-api-4ahn

// 1.
     const tenantId     = 'b5d500c1-e385-42eb-ab34-db10f116a192';
     const clientId     = '0d2192da-b16a-4f6b-8378-8cd1c08a799d';
// 2.
     const clientSecret = 'fIK8Q~Ef71EHU4DV7JAH4UqATC_7LEIb0wjIbdy5';
//   const clientSecret =  process.env.CLIENT_SECRET

     const clientConfig =
            {  auth:
                {  clientId
                ,  clientSecret
// 3.
                ,  authority: `https://login.microsoftonline.com/${ tenantId }`
                   }
               };
// 4.
     const authClient = new msal.ConfidentialClientApplication( clientConfig );

// -----------------------------------------------------------------

     const queryGraphApi = async ( path ) => {
// 5.
     const tokens    =  await authClient.acquireTokenByClientCredential(
// 6.
                            { scopes: ['https://graph.microsoft.com/.default']
                              } );

     const rawResult =  await fetch(`https://graph.microsoft.com/v1.0${ path }`
                          , { headers:
// 7.
                                { 'Authorization': `Bearer ${ tokens.accessToken }`
                                   }
                              } );

    return await rawResult.json( );
        }
// ------------------------------------------------------------------

    export { queryGraphApi };

// ------------------------------------------------------------------
// index.mjs
// ---------------

import { queryGraphApi } from './Auth3.mjs';

// In case you don't have top level await yet

  const sharePointSite = await queryGraphApi( '/sites/root' );
  console.log( sharePointSite );

/*
async function start() {
  const sharePointSite = await queryGraphApi('/sites/root');
  console.log(sharePointSite);
}

start().then(() => console.log('Complete.'));

*/
/*
sharePointSite
    code: 'AccessDenied'
    innerError: { date: '2023-06-29T03:08:19', request-id: '60bd311b-6dcc-44cd-b11f-25c9bb3d6594'
                , client-request-id: '60bd311b-6dcc-44cd-b11f-25c9bb3d6594'}
    message: 'Either scp or roles claim need to be present in the token.'
*/


