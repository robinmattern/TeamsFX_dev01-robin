
   import { queryGraphApi_v105 as queryGraphApi } from './Auth5.mjs';
   import { inspect       } from 'util'

//     var  aEntities = 'users'
       var  aEntities = 'notes1'
//     var  aEntities = 'notes2'

// -------  -----------------------------------------------------------------------------------------------------------------------

_Users: if (aEntities.match( /users/ ) ) {

//     var  pResult   =  await  queryGraphApi( '/users', 'User.Read.All' ); console.log( inspect( pResult.error, { depth: 9 } ) )
//     var  pResult   =  await  queryGraphApi( '/users' );                  console.log( inspect( pResult.value, { depth: 9 } ) )
//                       await         shoAPI( '/users', 'User.Read.All' )
//                       await         shoAPI( '/users', 'User.Read.All', '', shoUsers2 )
//                       await         shoAPI( '/users', 'User.Read.All', '', '',       "SicommNet,  Inc" )
//                       await         shoAPI( '/users', '',              '', '',       "SicommNet,  Inc" )
//                       await         shoAPI( '/users'  )
                         await         shoAPI( '/users', '',              '', shoUsers2 )
//                       await         shoAPI( '/users', '',              '', shoNotes2, '',             'v1.0' )

        if (1 == 2) {
            pResult = {
            error: {
                code:    'invalid_scope',
                message: '1002012 - [2023-06-30 22:00:46Z]: AADSTS1002012: The provided value for scope https://graph.microsoft.com/User.Read.All is not valid. Client credential flows must have a scope value with /.default suffixed to the resource identifier (application ID URI).\r\n' +
                         'Trace ID: ac0c7d3e-ee57-49d9-a3ce-e441ac0ee800\r\n' +
                         'Correlation ID: 3e9ed380-6cd8-4d4c-bed2-2f854b88271f\r\n' +
                         'Timestamp: 2023-06-30 22:00:46Z - Correlation ID: 3e9ed380-6cd8-4d4c-bed2-2f854b88271f - Trace ID: ac0c7d3e-ee57-49d9-a3ce-e441ac0ee800'
                          }
                }
            }
        };  // eif users
// -------  -----------------------------------------------------------------------------------------------------------------------

//_notes1: if (aEntities.match( /notes1/ )) { ... }
           if (aEntities.match( /notes1/ )) {

//     var  pResult   =  await  queryGraphApi(            '/onenote/notebooks', 'Notes.Read.All' ) // Error: invalid scope, use .default
       var  pResult   =  await  queryGraphApi(            '/onenote/notebooks' )                   // Token: only
//     var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks', '',                  'robin.mattern@sicomm.net' )

        if (pResult.error) {
            pResult.error.token = chop( 90, pResult.token, true )
            console.log( "\nError: \n" + inspect( pResult.error ) )
        } else {
       var  aToken    =  pResult.token; if (pResult.value) {
            console.log( shoNotes2( pResult.value ) ) }  // token only ??
            }
            console.log( `${ fmtDate(22).substring( 14 ) }  Complete. `.padEnd( 93, "-" ) + '\n')

//     var  pResult   =  await  queryGraphApi(            '/onenote/notebooks',  aToken )   // Token: only cuz no id
//     var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks',  aToken,      'robin.mattern@sicomm.net' )
       var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks', '',           'robin.mattern@sicomm.net' )  // works without token
       var  aToken    =  pResult.token; if (pResult.value) {
            console.log( shoNotes2( pResult.value ) ) }  // token only ??

                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken +'x', '5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6', shoNotes2, 'Bad Token' )
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken,      'e7c1f766-1963-4a4b-8296-aebceeb24fb1', shoNotes1, 'SicommNet Public Team' )
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken,      '5f7fabfc-9eea-4252-97f6-1b5b5___a0c6', shoNotes2, 'Bad Group' )

// Teams                id
// ------------------------------------   ------------------------------------- --------------------------------------
// 5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6   Admin
// 7faf391d-4632-480a-b73a-c1bef7ab1468   Robin's Team
// 6db02863-9c67-4de4-8e0b-5308b27bd8d8   Agency Support
// e7c1f766-1963-4a4b-8296-aebceeb24fb1   SicommNet Public Team
// d64997a6-72d3-499f-9d1f-8fc1bdcf66da   SicommNet Staff
// 3ff9068a-be48-4aa3-8f39-4853f72abb7b   Development
// -------------------------------------------------------------------------

//                       await         shoAPI( '/users', 'User.Read.All' )
//                       await         shoAPI( '/users/{id}/onenote/notebooks',  aToken, 'robin.mattern@sicomm.net',             shoNotes2 )

//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6', shoNotes2, 'Admin' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '6db02863-9c67-4de4-8e0b-5308b27bd8d8', shoNotes2, 'Agency Support' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '7faf391d-4632-480a-b73a-c1bef7ab1468', shoNotes2, 'Robin\'s Team' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'e7c1f766-1963-4a4b-8296-aebceeb24fb1', shoNotes2, 'SicommNet Public Team' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'd64997a6-72d3-499f-9d1f-8fc1bdcf66da', shoNotes2, 'SicommNet Staff' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '3ff9068a-be48-4aa3-8f39-4853f72abb7b', shoNotes2, 'Development' )

        };  // eif notes1
// -------  -----------------------------------------------------------------------------------------------------------------------

//_notes2: if (aEntities.match( /notes2/ )) {
        if (aEntities.match( /notes2/ )) {

//     var  pResult   =  await         getAPI( '/teams', '', shoTeams2, 'SicommNet' ); // if (pResult.error) { break _Notes2 }
       var  aToken    =  pResult.token

            pResult.data.map( async pTeam => {
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, pTeam.TeamId, shoNotes2, pTeam.Name )
                         } )

        };  // eif notes2
// -------  -----------------------------------------------------------------------------------------------------------------------

        if (aEntities.match( /messages/ ) ) {

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
        };  // eif messages
// -------  -----------------------------------------------------------------------------------------------------------------------

     async  function getAPI(  aPath, aScopes, aId, shoResult_, aIdName, aVersion ) {

       var  pResult  =  await queryGraphApi(  aPath, aScopes, aId, aVersion )

        if (pResult.error) {  return pResult }
        if (shoResult_) {
            pResult.data   =  shoResult_( pResult.value, true )
            }
     return pResult
            }  // eof getResult
// -------  -----------------------------------------------------------------------------------------------------------------------

     async  function shoAPI(  aPath, aScopes, aId, shoResult_, aIdName, aVersion ) {

       var  pResult  =  await queryGraphApi(  aPath, aScopes, aId, aVersion )
            console.log(  aIdName ? "  id:".padEnd(10) + aIdName + "\n" : "\n" )

        if (pResult.error) {  var aCR = "\n".padEnd(11)
            console.log( `Error: ** ${  pResult.error.code }${aCR}${ pResult.error.message.replace( /\n/g, aCR ) }` )
        } else {
        if (shoResult_) {
            pResult.data   =  shoResult_(  pResult.value , true ) }
            console.log(      shoResult(   pResult, shoResult_, aIdName ) )
            }
            console.log( `${  fmtDate(22).substring( 14 ) }  ${ aIdName ? `${aIdName} ` : "" }Complete. `.padEnd( 93, "-" ) + "\n" )

    return  pResult

//    ----  ---------------------------------------------

  function  shoResult( pJSON, shoJSON, aIdName ) {
        if (pJSON [   '@odata.context' ]) {   // API URL
        if (shoJSON) { var aFncName =  String( shoJSON ).match( /function +(.+)\(/ )[1] + "( pJSON.value )"
    return `Result:   ${aFncName}`   + shoJSON( pJSON.value,   aIdName )
         }  }
    return 'Result:   pJSON.value\n' + inspect( pJSON.value, { depth: 9 } )

            }  // eof shoResult
//    ----  ---------------------------------------------
         };  // eof shoAPI
// -------------------------------------------------------------------------

  function  shoNotes1( mNotebooks, aIdName ) {

       var  mRecs =    mNotebooks.map( pNotebook => {
       var  mFlds = [ "NotebookName: " + pNotebook.displayName
//                  , "New Date:     " + fmtDate( new Date, 8 )
//                  , "Today:        " + fmtDate( 10 )
//                  , "Null:         " + fmtDate('null' )
//                  , "MT:           " + fmtDate( 16, '' )
                    , "CreatedOn:    " + fmtDate( pNotebook.createdDateTime )
                    , "UpdatedOn:    " + fmtDate( 23, pNotebook.lastModifiedDateTime )  // 23 becomes 19, cuz date has only 19 digits: '2019-08-07T17:33:21Z'
                    , "CreatedBy:    " + pNotebook.createdBy.user.displayName
                    , "UpdatedBy:    " + pNotebook.lastModifiedBy.user.displayName
                    , "CreatedById:  " + pNotebook.createdBy.user.id
                    , "NotebookId:   " + pNotebook.id
                    , "WebPath:      " + pNotebook.links.oneNoteWebUrl.href
                    , "APIpath:      " + pNotebook.self
                      ]
    return  mFlds.join( "\n  " )
                  } )
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
       var  aRecs =  aLine + mRecs.join( aLine ) + aLine.replace( /\n$/, "" )
            aRecs = "\nNotebooks:" + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoNotes2( mNotebooks ) {

       var  mRecs =    mNotebooks.map( pNotebook => {
       var  mFlds = [  pNotebook.displayName.padEnd( 40 )
                    ,  fmtDate( pNotebook.lastModifiedDateTime )
                    ,  pNotebook.lastModifiedBy.user.displayName
                       ]
    return  mFlds.join( " " )
                  } )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nNotebooks:" + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoMsgs1(  mMessages ) {

       var  mRecs =    mMessages.map( pMessage => {
       var  mFlds = [ "Subject:          " +  pMessage.subject
                    , "From:             " +  pMessage.from.user.displayName
                    , "CreatedOn:        " +  fmtDate( 19, pMessage.createdDateTime        )     // 19 is the default
                    , "UpdatedOn:        " +  fmtDate(     pMessage.lastModifiedDateTime + '' )  // if you want 'null
                    , "DeletedOn:        " +  fmtDate( 22, pMessage.deletedDateTime      || '' ) // prevents fmtDate( 'now' )
                    , "Type:             " +  pMessage.messageType
                    , "BodyType:         " +  pMessage.body.contentType
                    , "BodyLength:       " +  pMessage.body.content.length
                    , "webUrl:           " +  pMessage.webUrl
                    , "AttachmentCnt:    " +  pMessage.attachments.length
                       ]
            pMessage.attachments.map( ( pAttachment, i ) => {
          mFlds.push( `AtchFile[${i+1}].Name: ` + pAttachment.name )
          mFlds.push( `AtchFile[${i+1}].Url : ` + pAttachment.contentUrl )
          mFlds.push( `AtchFile[${i+1}].Id  : ` + pAttachment.id )
                       } )
    return  mFlds.join( "\n  " )
                  } )
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
       var  aRecs = aLine + mRecs.join( aLine ) + aRecs + aLine.replace( /\n$/, "" )
            aRecs = "\nMessages:" + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoMsgs2(  mMessages ) {

       var  mRecs =    mMessages.map( ( pMessage , i ) => {
       var  aUser =               pMessage.from ? pMessage.from.user.displayName : "N/A"
       var  mFlds = [  fmtNum( 3, i + 1 ) + ". "
                    ,  chop(  30, aUser )
                    ,  fmtDate(   pMessage.lastModifiedDateTime ) + " "
                    ,  take(  18, pMessage.messageType )
                    ,  fmtNum( 6, pMessage.body.content ? pMessage.body.content.length : 0 )
                    ,  fmtBum( 3, pMessage.attachments  ? pMessage.attachments.length  : 0 ) + " "
//                  ,          `${pMessage.body.content ? pMessage.body.content.length : 0}`.padStart( 6 )
//                  ,          `${pMessage.attachments  ? pMessage.attachments.length  : 0}`.padStart( 3 ) + " "
                    ,  pMessage.subject
                       ]
    return  mFlds.join( " " )
                  } )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nMessages:" + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoUsers2( mUsers ) {

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
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nUsers:" + aRecs
    return  aRecs
            }
// -------------------------------------------------------------------------

  function  shoTeams2( mTeams, bData ) {

       var  mRecs =    mTeams.map( ( pRec , i ) => {
       var  mFlds = [  take( -3, i + 1 ) + "."
                    ,  chop( 40, pRec.displayName )
                    ,  chop( 30, pRec.mail )
                    ,  fmtPhone( pRec.mobilePhone ) + " "
                    ,            pRec.id
                    ,            pRec.userPrincipalName
                       ]
    return  bData ?  mFlds : mFlds.join( " " )
                  } )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nUsers:" + aRecs
    return  bData ?  mRecs : aRecs
            }
// -------------------------------------------------------------------------

  function  fmtPhone( a   ) {  a = a ? String(a).replace( /\+1 /, '' ).replace( /[-() ]/g, '' ) : ''  //
       if (!a) { return ''.padEnd(14) }
    return (a ? `(${ a.substring(0,3) }) ${ a.substring(3,6) }-${ a.substring(6,10) }` : '' ).padEnd(14)
            }
  function  fmtNum( n, m ) {
    return  take(  -n, m )
            }
  function  fmtDate( n, a ) {
        if (typeof(n) != 'number') { var x = a; a = n; n = x }; n = n ? n : 19
            a = ( typeof(a) != 'undefined' ) ? ( String(a).match( /Time/) ? a.toISOString() : a ) : new Date().toISOString();   // no typeof(a) == 'date'
    return  a.replace( /[TZ]/g, ' ' ).substring( (n == 8 ? 2 : 0), (n == 8 ? 10 : n ) )
            }
  function  take( n, a    ) {  a = String(  a )   // can be 'null'
    return (n < 0)        ? a.padStart( -n ) : a.padEnd( n )
            }
  function  chop( n, a, b ) {  a =  String( a )
            b = b ? `... (${a.length})` : '... '
    return (n < a.length) ? a.substring( 0, n - b.length ) + b : a.padEnd( n )
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

