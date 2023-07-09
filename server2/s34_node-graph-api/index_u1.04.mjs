
  import { queryGraphApi } from './Auth4.mjs';
  import { inspect       } from 'util'

// In case you don't have top level await yet

       var  aToken     =  'eyJ0eXAiOiJKV1QiLCJub25jZSI6InZSd1c2ZWhCTms2RVhXYkFEWWVNb2hILUhpUTd6MzdLbkdZa09yd0J2ek0iLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIvIiwiaWF0IjoxNjg4MDY0OTM1LCJuYmYiOjE2ODgwNjQ5MzUsImV4cCI6MTY4ODE1MTYzNSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhUQUFBQXl1U3VlQkxGUy9QVWtERGVYY0x4dTNjWXVCL2pFMVF1THNMcU1naEo0a2k2dWV3d3RrSzlVYTVvbDVPQmNMNDdQQzRRd3dudVhySmdPUFdINExZd3RtbnhKd2lzRlVpamlNZDZMU1YyS1dBPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiTWF0dGVybiIsImdpdmVuX25hbWUiOiJSb2JpbiIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjczLjEzMy44NC4xMDMiLCJuYW1lIjoiUm9iaW4gTWF0dGVybiIsIm9pZCI6ImIwMGY3ZDIxLWJhOTYtNDFkOC05YjIxLTg4NWIwNjc4ZDc0OSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzQkZGRDhDM0UyMEM4IiwicmgiOiIwLkFUY0F3UURWdFlYajYwS3JOTnNROFJhaGtnTUFBQUFBQUFBQXdBQUFBQUFBQUFBM0FPMC4iLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlByaXZpbGVnZWRPcGVyYXRpb25zLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BbGwgR3JvdXAuUmVhZFdyaXRlLkFsbCBJZGVudGl0eVJpc2tFdmVudC5SZWFkLkFsbCBNYWlsLlJlYWRXcml0ZSBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLlJlYWQgTm90ZXMuUmVhZC5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgUmVwb3J0cy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWQgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6Ijd6blpjdzFDRjgxa0JzZE9IN3Y5U0owY0tZZi1LdFdaWm5HZmthMzYzTVkiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiJiNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIiLCJ1bmlxdWVfbmFtZSI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInVwbiI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInV0aSI6IkMwMmphN3A1VlVtTTMtanZUcWxkQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfY2MiOlsiQ1AxIl0sInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6Ik1ENkFTb2J3ZXhTX0o4TmJIN0JJd29LdHpMTGtudzJxdUsxYzNEWlp4XzQifSwieG1zX3RjZHQiOjE0MTI4MDU0MDV9.LHQeYWxUaPn0K9U-87tsMYF9rweFanWAUBN0FSoerUIjzcBOk_WWVEFkT4iavKUEkTR0PqHNZ_6zFGguX5v_M4SXlwf5PDkDsiCxmiVfDyu_m0OcwGgyF4Uisu455kr5fZuGfXE9sDChdpE9_H3elNwrh3TUHn5IfXt5s7nJnsoWD8v2WNRwZZHC5WjTp7StwBc_IWds20jzb4YReIWqtT7GCZ5hllCQv94tedF9dSDKJLgj8m76ym1LehehnvfDNTRIxAhSTMDsj43y33bdBrX6SC9m8ZI0pAmr5QlOINHPFpcdkqozn7lRF0viZOUOwfBeXGY7fMOSEmvIOkZJ1A'
       var  aToken     =  'eyJ0eXAiOiJKV1QiLCJub25jZSI6IjVlX0Y0Z0stNlhVcWQ1eUQzaWpBeWstV00yckNlR0NmZW1xTjctZEFtWlkiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIvIiwiaWF0IjoxNjg4MTUxMjk3LCJuYmYiOjE2ODgxNTEyOTcsImV4cCI6MTY4ODIzNzk5NywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhUQUFBQTRNOVpLWjZ0QXBaQ1BzK2VPNzhUTkZ3R0lqbkxrWWNNWklrTlArUjZyakErWEVZejM3cU5PZXNOQVVEbG9jS2d4QUQyWUhsQ3hsWjRzeGZxbURiL0kxMWo4RHFXSVg4WmYyYUZwMHIrREp3PSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiTWF0dGVybiIsImdpdmVuX25hbWUiOiJSb2JpbiIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjczLjEzMy44NC4xMDMiLCJuYW1lIjoiUm9iaW4gTWF0dGVybiIsIm9pZCI6ImIwMGY3ZDIxLWJhOTYtNDFkOC05YjIxLTg4NWIwNjc4ZDc0OSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzQkZGRDhDM0UyMEM4IiwicmgiOiIwLkFUY0F3UURWdFlYajYwS3JOTnNROFJhaGtnTUFBQUFBQUFBQXdBQUFBQUFBQUFBM0FPMC4iLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlByaXZpbGVnZWRPcGVyYXRpb25zLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BbGwgR3JvdXAuUmVhZFdyaXRlLkFsbCBJZGVudGl0eVJpc2tFdmVudC5SZWFkLkFsbCBNYWlsLlJlYWRXcml0ZSBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLlJlYWQgTm90ZXMuUmVhZC5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgUmVwb3J0cy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWQgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6Ijd6blpjdzFDRjgxa0JzZE9IN3Y5U0owY0tZZi1LdFdaWm5HZmthMzYzTVkiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiJiNWQ1MDBjMS1lMzg1LTQyZWItYWIzNC1kYjEwZjExNmExOTIiLCJ1bmlxdWVfbmFtZSI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInVwbiI6InJvYmluLm1hdHRlcm5Ac2ljb21tLm5ldCIsInV0aSI6IlpLVVh1al9DLUVLOXhDOEtTcEN1QUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfY2MiOlsiQ1AxIl0sInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6Ik1ENkFTb2J3ZXhTX0o4TmJIN0JJd29LdHpMTGtudzJxdUsxYzNEWlp4XzQifSwieG1zX3RjZHQiOjE0MTI4MDU0MDV9.TLx5Pxv_OEglLlQmVx4LIoLTDdwmyAerxbPmXllXilPvHp3mv-LYfKZv71syaQ0kM2dMhj-v2UmN1rVYUxLcNQNKP7QumBNlpalXtxyTamw15AOeOtHDxzHPZcXU7TH5HwKAdHP8gmHoN0U3Wd0x3RYF2TzguTKC3nSvcoVUh2E1qOkt43coyxrgzB8fnt4f5LN4G0zd9nw4YYSth_dA_lduWVxx_hr0afkfBPo6OJ6OJT3ojZdwA-iWXLfF4_xt_0bDZV1-ioBTfPSxCDEpsiu6rwwlSwA7lGpWSb-MLtMOsFwt-3VHvtsuQxWwgYL8rkufWfVUhNzFNDRr9UxSPg'

//     var  pResult  =  await  queryGraphApi( '/sites/root' );
//     var  pResult  =  await  queryGraphApi( '/me', 'User.Read' )  // only valid for delegated, ie. interactive APIs

//     var  pResult  =  await  queryGraphApi( '/users', 'User.Read.All' )
//     var  pResult  =  await  queryGraphApi( '/users' ); console.log( inspect( pResult, {depth: 9 } ) )
                        await         shoAPI( '/users', '', '', shoUsers2 )
//                      await         shoAPI( '/users', '', '', shoNotes2, '', 'v1.0' )
//          
            process.exit()

if (1 == 2) {
            pResult = {
            error: {
                code: 'invalid_scope',
                message: '1002012 - [2023-06-30 22:00:46Z]: AADSTS1002012: The provided value for scope https://graph.microsoft.com/User.Read.All is not valid. Client credential flows must have a scope value with /.default suffixed to the resource identifier (application ID URI).\r\n' +
                'Trace ID: ac0c7d3e-ee57-49d9-a3ce-e441ac0ee800\r\n' +
                'Correlation ID: 3e9ed380-6cd8-4d4c-bed2-2f854b88271f\r\n' +
                'Timestamp: 2023-06-30 22:00:46Z - Correlation ID: 3e9ed380-6cd8-4d4c-bed2-2f854b88271f - Trace ID: ac0c7d3e-ee57-49d9-a3ce-e441ac0ee800'
                }
            }
}

//     var  pResult  =  await  queryGraphApi( '/onenote/notebooks', 'Notes.Read.All', 'robin.mattern@sicomm.net' )
//     var  pResult  =  await  queryGraphApi( '/onenote/notebooks',  aToken,          'robin.mattern@sicomm.net' )

// Teams                id
// ------------------------------------   ------------------------------------- --------------------------------------
// 5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6   Admin
// 7faf391d-4632-480a-b73a-c1bef7ab1468   Robin's Team
// 6db02863-9c67-4de4-8e0b-5308b27bd8d8   Agency Support
// e7c1f766-1963-4a4b-8296-aebceeb24fb1   SicommNet Public Team
// d64997a6-72d3-499f-9d1f-8fc1bdcf66da   SicommNet Staff
// 3ff9068a-be48-4aa3-8f39-4853f72abb7b   Development


// -------------------------------------------------------------------------

//   await  shoAPI( '/users/{id}/onenote/notebooks',  aToken, 'robin.mattern@sicomm.net',             shoNotes2 )

//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, '5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6', shoNotes2, 'Admin' )
//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, '6db02863-9c67-4de4-8e0b-5308b27bd8d8', shoNotes2, 'Agency Support' )
//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, '7faf391d-4632-480a-b73a-c1bef7ab1468', shoNotes2, 'Robin\'s Team' )
//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'e7c1f766-1963-4a4b-8296-aebceeb24fb1', shoNotes2, 'SicommNet Public Team' )
//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'd64997a6-72d3-499f-9d1f-8fc1bdcf66da', shoNotes2, 'SicommNet Staff' )
//   await  shoAPI( '/groups/{id}/onenote/notebooks', aToken, '3ff9068a-be48-4aa3-8f39-4853f72abb7b', shoNotes2, 'Development' )

// -------------------------------------------------------------------------

//  await  shoMsgs_4AgencySupport( )
           shoMsgs_4AgencySupport( shoMsgs2 )

// -------------------------------------------------------------------------

     async  function shoMsgs_4AgencySupport( shoMsgs ) {

       var  aTeam     = 'Agency Support'
       var  aTeamsId  = '6db02863-9c67-4de4-8e0b-5308b27bd8d8'
       var  mChannels =
                    [ [ '19:c2560595243948b1b57291798d19dc7f@thread.skype', ' General',                                'AgencySupport1@sicomm.net'         ]
                    , [ '19:f883b2ed4f6849cab9a68b08a1a9de59@thread.skype', ' - BASEC Issues',                         '329fcf0f.sicomm.net@amer.teams.ms' ]
                    , [ '19:021a5d821d0a4bcc827b39b4257c8bf0@thread.skype', ' - Corp Web Site',                        'db47a9d8.sicomm.net@amer.teams.ms' ]
                    , [ '19:a02390573eca470c92b3deb19ccc3a51@thread.skype', ' - Customer Issues, Hawaii',              '27412788.sicomm.net@amer.teams.ms' ]
                    , [ '19:67c94b1cf7f24302b19c6f91469a0f55@thread.skype', ' - Vendor Issues',                        'd16059cf.sicomm.net@amer.teams.ms' ]
                    , [ '19:d5a06be6d95044679402149897cb66e6@thread.skype', ' DevProjects',                            'fe7d106a.sicomm.net@amer.teams.ms' ]
                    , [ '19:ed856857800147249776262eac3c7cba@thread.skype', ' HAWAII - Training-Manuals-MMNew Issues', '137127f2.sicomm.net@amer.teams.ms' ]
                    , [ '19:6212b092aa384e278ed7db00779f26dd@thread.skype', ' iCatalog',                               '70ae15cb.sicomm.net@amer.teams.ms' ]
                    , [ '19:478beb18ddbf4676ae09c6e1f7a45762@thread.skype', ' Teams Rollout Project',                  '3c7e8a20.sicomm.net@amer.teams.ms' ]
                    , [ '19:a02f59f602fc419eb80fbce968dc325b@thread.skype', ' Website Configuration',                  'dd8194ce.sicomm.net@amer.teams.ms' ]
                        ]
//          -------------------------------------
//          https://graph.microsoft.com/beta/teams/{group-id-for-teams}/channels/{channel-id}/messages

      await mChannels.map( async function( mChannel )    {
//          mChannels.map( async         ( mChannel ) => { ... }
//          mChannels.map( async           mChannel   => { ... }
                   mChannel[1] = `${aTeam}, ${ mChannel[1].trim() }${ mChannel[2] ? '  ('+mChannel[2]+')' : '' }`
                        await shoAPI( `/teams/${aTeamsId}/channels/{id}/messages`, aToken, mChannel[0], shoMsgs, mChannel[1], 'beta' )
                   } )
//          -------------------------------------
            } // eof shoMsgs_4AgencySupport
// -------------------------------------------------------------------------

     async  function shoAPI(  aPath, aScopes, aId, shoResult_, aIdName, aVersion ) {

       var  pResult  =  await queryGraphApi(  aPath, aScopes, aId, aVersion )

        if (pResult.error) {
            console.log( ` ** ${ pResult.error.code }: ${ pResult.error.message }` )
        } else {
            console.log(   shoResult(  pResult, shoResult_, aIdName ) )
            }
            console.log( " Done."  + ( aIdName  ? "  " + aIdName : "" ) + "\n" )

//    ----  --------------------------------------------- 

  function  shoResult( pJSON, shoJSON, aIdName ) {
        if (pJSON[  '@odata.context' ]) {   // API URL
        if (shoJSON) {
    return  shoJSON( pJSON.value, aIdName ) }
            }
    return  inspect( pJSON, { depth: 9 } )
            }
//    ----  --------------------------------------------- 
            }
// -------------------------------------------------------------------------

  function  shoNotes1( mNotebooks, aIdName ) {

       var  mRecs =    mNotebooks.map( pNotebook => {
       var  mFlds = [ 'NotebookName: ' + pNotebook.displayName
                    , 'CreatedOn:    ' + pNotebook.createdDateTime.replace( /[TZ]/g, ' ' ).trim()
                    , 'UpdatedOn:    ' + pNotebook.lastModifiedDateTime.replace( /[TZ]/g, ' ' ).trim()
                    , 'CreatedBy:    ' + pNotebook.createdBy.user.displayName
                    , 'UpdatedBy:    ' + pNotebook.lastModifiedBy.user.displayName
                    , 'CreatedById:  ' + pNotebook.createdBy.user.id
                    , 'NotebookId:   ' + pNotebook.id
                    , 'WebPath:      ' + pNotebook.links.oneNoteWebUrl.href
                    , 'APIpath:      ' + pNotebook.self
                      ]
    return  mFlds.join( '\n' )
              } )

       var  aLine = '\n'.padEnd( 50, '-' ) + '\n'
       var  aRecs = aLine + mRecs.join( aLine )

            aRecs = '\nNotebooks:' + aRecs
            aRecs = (aIdName ? 'id: ' + aIdName + '\n' : '' ) + aRecs
    return  aRecs
              }
// -------------------------------------------------------------------------

  function  shoNotes2( mNotebooks, aIdName ) {

       var  mRecs =    mNotebooks.map( pNotebook => {
       var  mFlds = [  pNotebook.displayName.padEnd( 40 )
                    ,  pNotebook.lastModifiedDateTime.replace( /[TZ]/g, ' ' )
                    ,  pNotebook.lastModifiedBy.user.displayName
                       ]
    return  mFlds.join( ' ' )
              } )
       var  aLine = '\n '
       var  aRecs = aLine + mRecs.join( aLine )

            aRecs = '\nNotebooks:' + aRecs
            aRecs = (aIdName ? ' id:    ' + aIdName + '\n' : '' ) + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoMsgs1(  mMessages, aIdName ) {

       var  mRecs =    mMessages.map( pMessage => {
       var  mFlds = [ 'Subject:          ' +  pMessage.subject
                    , 'From:             ' +  pMessage.from.user.displayName
                    , 'CreatedOn:        ' +( pMessage.createdDateTime      || '' ).replace( /[TZ]/g, ' ' ).trim()
                    , 'UpdatedOn:        ' +( pMessage.lastModifiedDateTime || '' ).replace( /[TZ]/g, ' ' ).trim()
                    , 'DeletedOn:        ' +( pMessage.deletedDateTime      || '' ).replace( /[TZ]/g, ' ' ).trim()
                    , 'Type:             ' +  pMessage.messageType
                    , 'BodyType:         ' +  pMessage.body.contentType
                    , 'BodyLength:       ' +  pMessage.body.content.length
                    , 'webUrl:           ' +  pMessage.webUrl
                    , 'AttachmentCnt:    ' +  pMessage.attachments.length
                       ]
            pMessage.attachments.map( ( pAttachment, i ) => {
          mFlds.push( `AtchFile[${i+1}].Name: ` + pAttachment.name )
          mFlds.push( `AtchFile[${i+1}].Url : ` + pAttachment.contentUrl )
          mFlds.push( `AtchFile[${i+1}].Id  : ` + pAttachment.id )
                       } )
    return  mFlds.join( '\n' )
              } )

       var  aLine = '\n'.padEnd( 50, '-' ) + '\n'
       var  aRecs = aLine + mRecs.join( aLine )

            aRecs = '\nMessages:' + aRecs
            aRecs = (aIdName ? 'id: ' + aIdName + '\n' : '' ) + aRecs
    return  aRecs
              }
// -------------------------------------------------------------------------

  function  shoMsgs2(  mMessages, aIdName ) {

       var  mRecs =    mMessages.map( ( pMessage , i ) => {
       var  aUser =    pMessage.from ? pMessage.from.user.displayName : 'N/A'
       var  mFlds = [ (i+1).toString().padStart(4) + ". "
                    ,  chop( 30, aUser )
                    ,  pMessage.lastModifiedDateTime.replace( /[TZ]/g, ' ' ).substring( 0, 19 ) + " "
                    ,  pMessage.messageType.padEnd( 18 )
                   ,`${pMessage.body.content ? pMessage.body.content.length : 0}`.padStart(6)
                   ,`${pMessage.attachments  ? pMessage.attachments.length  : 0}`.padStart(3) + " "
                    ,  pMessage.subject
                       ]
    return  mFlds.join( ' ' )
                       } )

       var  aLine = '\n '
       var  aRecs = aLine + mRecs.join( aLine )

            aRecs = '\nMessages:' + aRecs
            aRecs = (aIdName ? ' id:    ' + aIdName + '\n' : '' ) + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoUsers2(  mUsers, aIdName ) {

       var  mRecs =    mUsers.map( ( pRec , i ) => {
       var  mFlds = [  take( -3, i + 1 ) + "."
                    ,  chop( 40, pRec.displayName )
                    ,  chop( 30, pRec.mail )
                    ,  fmtPhone( pRec.mobilePhone ) + " "
                    ,            pRec.id
                    ,            pRec.userPrincipalName
                       ]
    return  mFlds.join( " " )
                       } )

       var  aLine = "\n"
       var  aRecs =  aLine + mRecs.join( aLine )

            aRecs = "\nUsers:" + aRecs
            aRecs =( aIdName ? " id:    " + aIdName + "\n" : "" ) + aRecs
    return  aRecs
            }

// -------------------------------------------------------------------------

  function  fmtPhone( a ) { a = a ? String(a).replace( /\+1 /, '' ).replace( /[-() ]/g, '' ) : ''  // 
       if (!a) { return ''.padEnd(14) }
    return (a ? `(${ a.substring(0,3) }) ${ a.substring(3,6) }-${ a.substring(6,10) }` : '' ).padEnd(14)    
            }
  function  take( n, a ) {  a = String(  a )   // can be 'null' 
    return (n < 0)        ? a.padStart( -n ) : a.padEnd( n ) 
            }
  function  chop( n, a ) {  a =  String( a )
    return (n < a.length) ? a.substring( 0, n-4 ) + '... ' : a.padEnd( n )
            }
// -------------------------------------------------------------------------

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

