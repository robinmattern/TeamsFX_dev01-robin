
   import { queryGraphApi_v105 as queryGraphApi } from './Auth5.mjs';
   import { getAPI, shoAPI } from './Auth5.mjs';
   import { inspect        } from 'util'

//     var  aEntities = 'users'
//     var  aEntities = 'notes1'
       var  aEntities = 'notes2'

// -------  -----------------------------------------------------------------------------------------------------------------------

_Users: if (aEntities.match( /users/ ) ) {

//     var  pResult   =  await  queryGraphApi( '/users', 'User.Read.All' ); console.log( inspect( pResult.error, { depth: 9 } ) )
//     var  pResult   =  await  queryGraphApi( '/users' );                  console.log( inspect( pResult.value, { depth: 9 } ) )
//                       await         shoAPI( '/users', 'User.Read.All' )
//                       await         shoAPI( '/users', 'User.Read.All', '', fmtUsers2 )
//                       await         shoAPI( '/users', 'User.Read.All', '', '',       "SicommNet,  Inc" )
//                       await         shoAPI( '/users', '',              '', '',       "SicommNet,  Inc" )
//                       await         shoAPI( '/users'  )
                         await         shoAPI( '/users', '',              '', fmtUsers2 )
//                       await         shoAPI( '/users', '',              '', fmtNotes2, '',             'v1.0' )

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

_notes1: if (aEntities.match( /notes1/ )) {

//     var  pResult   =  await  queryGraphApi(            '/onenote/notebooks', 'Notes.Read.All' ) // Error: invalid scope, use .default
       var  pResult   =  await  queryGraphApi(            '/onenote/notebooks' )                   // Token: only
//     var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks', '',                  'robin.mattern@sicomm.net' )

        if (pResult.error) {
            pResult.error.token = chop( 90, pResult.token, true )
            console.log( "\nError: \n" + inspect( pResult.error ) )
        } else {
       var  aToken    =  pResult.token; if (pResult.value) {
            console.log( fmtNotes2( pResult.value ) ) }  // token only ??
            }
            console.log( `${ fmtDate(22).substring( 14 ) }  Complete. `.padEnd( 87, "-" ) + '\n')

//     var  pResult   =  await  queryGraphApi(            '/onenote/notebooks',  aToken )   // Token: only cuz no id
//     var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks',  aToken,      'robin.mattern@sicomm.net' )
       var  pResult   =  await  queryGraphApi( '/users/{id}/onenote/notebooks', '',           'robin.mattern@sicomm.net' )  // works without token
       var  aToken    =  pResult.token; if (pResult.value) {
            console.log( fmtNotes2( pResult.value ) ) }  // token only ??

                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken +'x',  fmtNotes2, '5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6','Bad Token' )
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken,       fmtNotes1, 'e7c1f766-1963-4a4b-8296-aebceeb24fb1','SicommNet Public Team' )
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken,       fmtNotes2, '5f7fabfc-9eea-4252-97f6-1b5b5___a0c6','Bad Group' )

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
//                       await         shoAPI( '/users/{id}/onenote/notebooks',  aToken, 'robin.mattern@sicomm.net',             fmtNotes2 )

//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '5f7fabfc-9eea-4252-97f6-1b5b5de9a0c6', fmtNotes2, 'Admin' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '6db02863-9c67-4de4-8e0b-5308b27bd8d8', fmtNotes2, 'Agency Support' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '7faf391d-4632-480a-b73a-c1bef7ab1468', fmtNotes2, 'Robin\'s Team' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'e7c1f766-1963-4a4b-8296-aebceeb24fb1', fmtNotes2, 'SicommNet Public Team' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, 'd64997a6-72d3-499f-9d1f-8fc1bdcf66da', fmtNotes2, 'SicommNet Staff' )
//                       await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, '3ff9068a-be48-4aa3-8f39-4853f72abb7b', fmtNotes2, 'Development' )

        };  // eif notes1
// -------  -----------------------------------------------------------------------------------------------------------------------

//function  selectTeam3( pTeam ) { return  pTeam.TeamName.match( /Support|Development/ ) }

_notes2: if (aEntities.match( /notes2/ )) {

  function  selectTeam3( pTeam ) { return  pTeam.TeamName.match( /Support|Development/ ) }   // not defined globally; where( 'selectTeam3' ) won't work 

//     var  pResult   =  await         shoAPI( '/joinedTeams',            '', fmtTeams1, '',                         'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams2, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams3, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
       var  pResult   =  await         getAPI( '/users/{id}/joinedTeams', '', fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );

        if (pResult.error) { break _notes2 }
       var  aToken    =  pResult.token

            pResult.data.map( async pTeam => {  // crazy!!  S.B. async function( pTeam ) { ...  }
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, fmtNotes3, pTeam.Id, pTeam.TeamName )
                         } )

        };  // eif notes2

// API requires one of  'Team.ReadBasic.All, TeamSettings.Read.All, TeamSettings.ReadWrite.All'
// Roles on the request     'Notes.Read.All,         User.Read.All,        Notes.ReadWrite.All'.

// -------  -----------------------------------------------------------------------------------------------------------------------

_messages1: if (aEntities.match( /messages1/ ) ) {

//  await  shoMsgs_4AgencySupport( )
           shoMsgs_4AgencySupport( shoMsgs2 )

// -------------------------------------------------------------------------

     async  function shoMsgs_4AgencySupport( fmtMsgs ) {

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
                        await shoAPI( `/teams/${aTeamsId}/channels/{id}/messages`, aToken, mChannel[0], fmtMsgs, mChannel[1], 'beta' )
                   } )
//          -------------------------------------
            } // eof shoMsgs_4AgencySupport
// -------------------------------------------------------------------------
        };  // eif messages
// -------  -----------------------------------------------------------------------------------------------------------------------

_messages2: if (aEntities.match( /messages2/ )) {

       var  pResult   =  await         getAPI( '/users/{id}/joinedTeams', '', fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );

        if (pResult.error) { break _messages2 }
       var  aToken    =  pResult.token

            pResult.data.map( async pTeam => {  // crazy!!  S.B. async function( pTeam ) { ...  }
                         await         shoAPI( '/groups/{id}/channels', aToken, fmtChannel3, pTeam.Id, pTeam.TeamName )
                         } )

        };  // eif messages2























// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtNotes1( mData, asData ) { return fmtNotes_1( mData, asData ) }
  function  fmtNotes2( mData, asData ) { return fmtNotes_2( mData, asData ) }
  function  fmtNotes3( mData, asData ) { return fmtNotes_1( mData,'byRows') }

  function  fmtNotes_1( mNotebooks, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'asData'

       var  mRecs =    mNotebooks.map( pRec => {
       var  pFlds = { "NotebookName": chop( 40, pRec.displayName )
                    , "UpdatedOn"   : fmtDate(  pRec.lastModifiedDateTime ) + " " 
                    , "UpdatedBy"   :           pRec.lastModifiedBy.user.displayName
                       }
    return  pFlds
            } )
//     var  mData = fmtData( mRecs, aData, "Teams", 'qq' )  // bQQ  if bCSV == false
//     var  mData = fmtData( mRecs, aData, "Teams", true )  // bCSV if asData != 'asData'
       var  mData = fmtData( mRecs, aData, "Teams" )  // bCSV if asData != 'asData'
    return  mData
         }; // eof fmtNotes_1
// -------  ------------------------------------------------------------------

  function  fmtNotes_2( mNotebooks, asData ) { // byFlds
       var  aData =    asData ? asData : 'byFlds', bData = (aData == 'asData')

       var  mRecs =    mNotebooks.map( pRec => {
       var  mFlds = [ "NotebookName: " +          pRec.displayName
//                  , "New Date:     " + fmtDate( new Date, 8 )
//                  , "Today:        " + fmtDate( 10 )
//                  , "Null:         " + fmtDate('null' )
//                  , "MT:           " + fmtDate( 16, '' )
                    , "CreatedOn:    " + fmtDate( pRec.createdDateTime )
                    , "UpdatedOn:    " + fmtDate( 23, pRec.lastModifiedDateTime )  // 23 becomes 19, cuz date has only 19 digits: '2019-08-07T17:33:21Z'
                    , "CreatedBy:    " +          pRec.createdBy.user.displayName
                    , "UpdatedBy:    " +          pRec.lastModifiedBy.user.displayName
                    , "CreatedById:  " +          pRec.createdBy.user.id
                    , "NotebookId:   " +          pRec.id
                    , "WebPath:      " +          pRec.links.oneNoteWebUrl.href
                    , "APIpath:      " +          pRec.self
                      ]
    return  bData ? mFlds : mFlds.join( "\n  " ) // Should be pFlds
                  } )
        if (bData) { return  mRecs } else {
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
       var  aRecs =  aLine + mRecs.join( aLine ) + aLine.replace( /[\n ]+$/, "" )
            aRecs = "\nNotebooks:" + aRecs
    return  aRecs }
         }; // eof fmtNotes_2
// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtUsers1( mData ) { return fmtUsers_3( mData, 'asData' ) }
  function  fmtUsers2( mData ) { return fmtUsers_3( mData, 'byFlds' ) }
  function  fmtUsers3( mData ) { return fmtUsers_3( mData, 'byRows' ) }

  function  fmtUsers_3( mUsers, asData ) { // asData, byFlds or byRows

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

         }; // eof fmtUsers_2
// -------  -----------------------------------------------------------------------------------------------------------------------













// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtMsgs1( mData ) { return fmtMsgs_1( mData, 'asData' ) }
  function  fmtMsgs2( mData ) { return fmtMsgs_2( mData, 'byFlds' ) }
  function  fmtMsgs3( mData ) { return fmtMsgs_1( mData, 'byRows' ) }

  function  fmtMsgs_1(  mMessages ) {  // asData or byRows

       var  mRecs =    mMessages.map( ( pRec , i ) => {
       var  aUser =               pRec.from ? pRec.from.user.displayName : "N/A"
       var  mFlds = [  fmtNum( 3, i + 1 ) + ". "
                    ,  chop(  30, aUser )
                    ,  fmtDate(   pRec.lastModifiedDateTime ) + " "
                    ,  take(  18, pRec.messageType )
                    ,  fmtNum( 6, pRec.body.content ? pRec.body.content.length : 0 )
                    ,  fmtBum( 3, pRec.attachments  ? pRec.attachments.length  : 0 ) + " "
                    ,  pRec.subject
                       ]
    return  mFlds.join( " " )
                  } )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nMessages:" + aRecs
    return  aRecs

         }; // eof fmtMsgs_1
// -------  ------------------------------------------------------------------

  function  fmtMsgs_2( mMessages, asData ) {  // byFlds

       var  mRecs =    mMessages.map( pRec => {
       var  mFlds = [ "Subject:       " +               pRec.subject
                    , "From:          " +               pRec.from.user.displayName
                    , "CreatedOn:     " +  fmtDate( 19, pRec.createdDateTime        )     // 19 is the default
                    , "UpdatedOn:     " +  fmtDate(     pRec.lastModifiedDateTime + '' )  // if you want 'null
                    , "DeletedOn:     " +  fmtDate( 22, pRec.deletedDateTime      || '' ) // prevents fmtDate( 'now' )
                    , "Type:          " +               pRec.messageType
                    , "BodyType:      " +               pRec.body.contentType
                    , "BodyLength:    " +               pRec.body.content.length
                    , "WebUrl:        " +               pRec.webUrl
                    , "AttachmentCnt: " +               pRec.attachments.length
                       ]
            pRec.attachments.map( ( pAttachment, i ) => {
          mFlds.push( `AtchFile[${i+1}].Name: ` + pAttachment.name )
          mFlds.push( `AtchFile[${i+1}].Url : ` + pAttachment.contentUrl )
          mFlds.push( `AtchFile[${i+1}].Id  : ` + pAttachment.id )
                       } )
    return  mFlds.join( "\n  " )
                  } )
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
//     var  aRecs = aLine + mRecs.join( aLine ) + aRecs + aLine.replace( /[\n ]+$/, "" )
       var  aRecs = aLine + mRecs.join( aLine ) +         aLine.replace( /[\n ]+$/, "" )
            aRecs = "\nMessages:" + aRecs
    return  aRecs
         }; // eof fmtMsgs_2
// -------  -----------------------------------------------------------------------------------------------------------------------

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
// -------  ------------------------------------------------------------------





// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtTeams1( mData, asData ) { return fmtTeams_1( mData, asData ) }
  function  fmtTeams2( mData, asData ) { return fmtTeams_2( mData, asData ) }
//function  fmtTeams3( mData, asData ) { return fmtTeams_3( mData,'byRows') }  // Sometimes we want to return data
  function  fmtTeams3( mData, asData ) { return fmtTeams_3( mData, asData ) }

  function  fmtTeams( mData, nFmt, asData ) {
          switch (nFmt ? nFmt : 2) {
            case 1: fmtTeams_1( mData, asData ? asData : 'asData' ); break
            case 2: fmtTeams_2( mData, asData ? asData : 'byFlds' ); break
            case 3: fmtTeams_3( mData, asData ? asData : 'byRows' ); break
         }  }

  function  fmtTeams_1( mTeams, asData ) {  // asData or byRows
       var  aData =    asData ? asData : 'asData'
        if (typeof(selectTeam1) != 'function') { selectTeams1 = () => true } // avoids error if not defined globally 

       var  mRecs =    mTeams.map( ( pRec , i ) => {
       var  pFlds = { 'TeamName'   :          pRec.displayName
                    , 'Id'         :          pRec.id
                    , 'Description':          pRec.description
                    , 'CreatedOn'  : fmtDate( pRec.createdDateTime + '' ).trim()
                    , 'WebUrl'     :          pRec.webUrl
                       }
    return  pFlds
            } )
//     var  mData =   fmtData( mRecs, aData, "Teams", true )  // bCSV
            mRecs =   mRecs.filter( selectTeam1 )
       var  mData =   fmtData( mRecs, aData, "Teams" )        // aData could be 'asCSV'
    return  mData

         }; // eof fmtTeams_1
// -------  ------------------------------------------------------------------

  function  fmtTeams_2( mTeams, asData ) {  // byFlds
       var  aData =    asData ? asData : 'byFlds', bData = (aData == 'asData')

       var  mRecs =    mTeams.map( ( pRec , i ) => {
       var  pFlds = { 'No'         : take( -3, i + 1 ) + "."
                    , 'TeamName'   : chop( 25, pRec.displayName )
                    , 'Id'         :           pRec.id
                    , 'CreatedOn'  : fmtDate(  pRec.createdDateTime + '' )
                    , 'WebUrl'     : chop( 35, pRec.webUrl )
                    , 'Description':           pRec.description
                       }
        if (bData) { return  pFlds } else {
       var  mFlds =  Object.keys( pFlds ).map( aFld => pFlds[ aFld ] )
    return  mFlds.join( "\n  " )
            } } )
        if (bData) { return  mRecs } else {
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
       var  aRecs =  aLine + mRecs.join( aLine ) + aLine.replace( /[\n ]+$/, "" )
            aRecs = "\nTeams:" + aRecs
    return  aRecs
            }
         }; // eof fmtTeans_2
// -------  ------------------------------------------------------------------

/*function  fmtTeams_3( mTeams, asData ) {  // asData or byRows
       var  aData =    asData ? asData : 'byRows'
    return  fmtTeams_2( mTeams, asData )  if fmtTeams_2 was smart
            }
 */
  function  fmtTeams_3( mTeams, asData ) {  // byRows
       var  aData =    asData ? asData : 'byRows'

       var  mRecs =    mTeams.map( ( pRec , i ) => {
       var  pFlds = { 'No'         : take( -3, i + 1 ) + "."
                    , 'TeamName'   : chop( 25, pRec.displayName )
                    , 'Id'         :           pRec.id + " "
//                  , 'CreatedOn'  : fmtDate(  pRec.createdDateTime + '' )  // always null
//                  , 'WebUrl'     : chop( 35, pRec.webUrl )                // always null
                    , 'Description': chop( 70, pRec.description )
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter( where( 'selectTeam3' ) )  // must be global
            mRecs =   mRecs.filter(           selectTeam1    )  // must be global and defined
//     var  mData = fmtData( mRecs, aData, "Teams", "qq" )
       var  mData = fmtData( mRecs, aData, "Teams" )
    return  mData

         }; // eof fmtTeams_3
// -------  ------------------------------------------------------------------

  function  selectTeam1(  pTeam  ) { 
    return  pTeam.TeamName.match( /Support|Development/ ) 
            } 

// -------  ------------------------------------------------------------------

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
    return  a.replace( /[TZ]/g, ' ' ).padEnd( n ).substring( (n == 8 ? 2 : 0), (n == 8 ? 10 : n ) )
            }
  function  take( n, a    ) {  a = String(  a )   // can be 'null'
    return (n < 0)        ? a.padStart( -n ) : a.padEnd( n )
            }
  function  chop( n, a, b ) {  a =  String( a )
            b = b ? `... (${a.length})` : '... '
    return (n < a.length) ? a.substring( 0, n - b.length ) + b : a.padEnd( n )
            }
  function  where( f ) { var b = eval( `typeof( this.${f} ) == 'function'` ) ; 
//   return b ? eval( f ) : () => true }            
     return b ? this[f] : () => true }            
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

