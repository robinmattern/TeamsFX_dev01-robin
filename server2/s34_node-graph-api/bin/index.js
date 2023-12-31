#!/usr/bin/env node

// read in env settings
    require( 'dotenv' ).config( );

     const  yargs   =  require( 'yargs'   );

     const  fetch   =  require( './fetch' );
     const  auth    =  require( './auth'  );

     const  options =  yargs.usage(  'Usage: --op <operation_name>' )
                            .option( 'op', { alias: 'operation', describe: 'operation name', type: 'string', demandOption: true } )
                            .argv;

//--------------------------------------------------

     async  function main( ) {
            console.log(`You have selected: ${options.op}`);

    switch (yargs.argv['op']) {

        case 'getUsers':

            try { // here we get an access token
                
                const authResponse  =  await auth.getToken( auth.tokenRequest );
                const users         =  await fetch.callApi( auth.apiConfig.uri, authResponse.accessToken ); // call the web API with the access token
                
                console.log( "  Users:", users ); // display result

            } catch (error) {
                console.log( "  Error:", error );
                }
                break;

        default:
                console.log('Select a Graph operation first');
                break;
        }   
    };
//--------------------------------------------------

    main( );
    