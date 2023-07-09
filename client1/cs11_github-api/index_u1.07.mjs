   import { queryGraphApi_v105 as queryGraphApi } from './Auth5.mjs';
   import { getAPI, shoAPI, fmtData } from './Auth5.mjs';
   import { inspect                 } from 'util'

//     var  aEntities = 'teams1'
//     var  aEntities = 'users1'
//     var  aEntities = 'notes1'
//     var  aEntities = 'notes2'
       var  aEntities = 'messages2'
//     var  aEntities = 'threads1'
//     var  aEntities = 'groups1'

// -------  -----------------------------------------------------------------------------------------------------------------------

if (aEntities.match( /teams1/ )) {

               getTeams1( )  

async function getTeams1( ) { 
       var  aTeams_API    =  '/users/{id}/joinedTeams'
       var  aScopes1      =  ''

//    ----  --------------------------------------------------------

//     var  pTeams    =  await  getAPI( aTeams_API, aScopes1, fmtTeams3, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );
       var  pTeams    =  await  getAPI( aTeams_API, aScopes1, fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );
//     var  pTeams    =  await  shoAPI( aTeams_API, aScopes1, fmtTeams2, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );
        if (pTeams.error) {     return } 

            } // eof getTeams1
// -------------------------------------------------------------------------
        };  // eif teams1
// -------  -----------------------------------------------------------------------------------------------------------------------

_Users: if (aEntities.match( /users1/ ) ) {

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
            }   }         }
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

        };  // eif notes1
// -------  -----------------------------------------------------------------------------------------------------------------------

_notes2: if (aEntities.match( /notes2/ )) {

  function  selectTeam3( pTeam ) { return  pTeam.Name.match( /Support|Development/ ) }   // not defined globally; where( 'selectTeam3' ) won't work

//     var  pResult   =  await         shoAPI( '/joinedTeams',            '', fmtTeams1, '',                         'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams2, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
//     var  pResult   =  await         shoAPI( '/users/{id}/joinedTeams', '', fmtTeams3, 'robin.mattern@sicomm.net', 'SicommNet, Inc' ); // Resource not found for the segment 'joinedTeams'.
       var  pResult   =  await         getAPI( '/users/{id}/joinedTeams', '', fmtTeams1, 'robin.mattern@sicomm.net', 'SicommNet, Inc' );

        if (pResult.error) { break _notes2 }
       var  aToken    =  pResult.token

            pResult.data.map( async pTeam => {  // crazy!!  S.B. async function( pTeam ) { ...  }
                         await         shoAPI( '/groups/{id}/onenote/notebooks', aToken, fmtNotes3, pTeam.Id, pTeam.Name )
                         } )

        };  // eif notes2

// API requires one of  'Team.ReadBasic.All, TeamSettings.Read.All, TeamSettings.ReadWrite.All'
// Roles on the request     'Notes.Read.All,         User.Read.All,        Notes.ReadWrite.All'.

// -------  -----------------------------------------------------------------------------------------------------------------------
// -------  -----------------------------------------------------------------------------------------------------------------------

_messages1: if (aEntities.match( /messages1/ ) ) {

//  await  shoMsgs_4AgencySupport( )
           shoMsgs_4AgencySupport( shoMsgs2 )

// -------------------------------------------------------------------------

     async  function shoMsgs_4AgencySupport( fmtMessages ) {

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
                        await shoAPI( `/teams/${ aTeamsId }/channels/{id}/messages`, aToken, mChannel[0], fmtMessages, mChannel[1], 'beta' )
                   } )
//          -------------------------------------
            } // eof shoMsgs_4AgencySupport
// -------------------------------------------------------------------------
        };  // eif messages
// -------  -----------------------------------------------------------------------------------------------------------------------

_messages2: if (aEntities.match( /messages2/ )) {

                getMessages2() 

//    --------------------------------------------------------------------

 async function getMessages2( ) { 
       var  aTeams_API    =  '/users/{id}/joinedTeams'
       var  aChannels_API =  '/teams/{id}/channels'
       var  aMessages_API =  '/teams/{TeamId}/channels/{id}/messages'
       var  aScopes1      =  ''
//     var  aScopes2      =  'Channel.ReadBasic.All'

//    ----  --------------------------------------------------------

//     var  pTeams    =  await  getAPI( aTeams_API,    aScopes1, fmtTeams1,    'robin.mattern@sicomm.net', 'SicommNet, Inc' );
       var  pTeams    =  await  shoAPI( aTeams_API,    aScopes1, fmtTeams3,    'robin.mattern@sicomm.net', 'SicommNet, Inc' );
        if (pTeams.error) {     return } 

       var  aToken    =  pTeams.token 
//    ----  --------------------------------------------------------

//          pTeams.data.map(          function( pTeam )    { // No workie!!  S.B. async function( pTeam ) { ...  } ... }
//          pTeams.data.map(    async function( pTeam )    { // Not crazy!!  S.B. async function( pTeam ) { ...  } ... }
//          pTeams.data.map(    async         ( pTeam ) => { // crazy!!  S.B. async function( pTeam ) { ...  }
            pTeams.data.map(    async           pTeam   => { // crazy!!  S.B. async function( pTeam ) { ...  }

       var  pChannels =  await  shoAPI( aChannels_API, aToken,   fmtChannels3,  pTeam.Id,    pTeam.Name )
        if (pChannels.error) {  return } 

//          -----------------------------------------------------

            pChannels.data.map( async   pChannel => {

       var  aChannel  = `${pTeam.Name.trim()}, ${ pChannel.Name.trim() }${ pChannel.mail ? '  (' + pChannel.mail.trim() + ')' : '' }`
       var  aMsgs_API =  aMessages_API.replace( /{TeamId}/, pTeam.Id.trim() )

//          -----------------------------------------------

       var  pMessages =  await  shoAPI( aMsgs_API,     aToken,   fmtMessages3,  pChannel.Id, aChannel, 'beta' )
        if (pMessages.error) {  return }

            } ) // eol mChannels
//          -----------------------------------------------------
            } ) // eol mTeams
//    ----  --------------------------------------------------------
      }; // eof getMessages2  
//    --------------------------------------------------------------------
    };                   // eif messages2
// -------  -----------------------------------------------------------------------------------------------------------------------

       if (aEntities.match( /threads1/ )) {

                getThreads1() 

//    --------------------------------------------------------------------

 async function getThreads1( ) { 

       var  pAPIs = { }

            pAPIs.Teams    = { API:  '/users/{userId}/joinedTeams'
                             , Name: '{userName} teams' }

            pAPIs.Channels = { API:  '/teams/{teamId}/channels'
                             , Name: '{teamsName} channels' }

            pAPIs.Messages = { API:  '/teams/{teamId}/channels/{channelId}/messages'
                             , Name: '{teamName} - {channeName} messages' }

            pAPIs.Threads1 = { API:  '/groups/{groupId}/threads'
                             , Name: '{groupName} threads' }

            pAPIs.Convers  = { API:  '/groups/{groupId}/conversations'
                             , Name: '{groupName} conversations' }

            pAPIs.Posts1   = { API:  '/groups/{groupId}/threads/{threadId}/posts' 
                             , Name: '{groupName} - {threadName} posts' }

            pAPIs.Threads2 = {_API:  '/groups/{groupId}/conversations/{conversationId}/threads'
                             , Name: '{groupName} - {conversationName} threads' }

            pAPIs.Posts2   = { API:  '/groups/{groupId}/conversations/{conversationId}/threads/{threadId}/posts' 
                             , Name: '{groupName} - {threadName} posts' }

       var  aScopes1       =  ''
//     var  aScopes2       =  'Channel.ReadBasic.All'

       var  aUserId        =  'robin.mattern@sicomm.net',                         aUserName =  aUserId 

       var  aTeamId        =  'e7c1f766-1963-4a4b-8296-aebceeb24fb1',             aTeamName = "SicommNet Public Team" 
       var  aTeamId        =  '6db02863-9c67-4de4-8e0b-5308b27bd8d8',             aTeamName = "Agency Support" 

       var  aChannelId     =  '19:f883b2ed4f6849cab9a68b08a1a9de59@thread.skype', aChannelName  = "- BASEC Issues"                
       var  aChannelId     =  '19:a02390573eca470c92b3deb19ccc3a51@thread.skype', aChannelName  = "- Customer Issues, Hawaii"     
       var  aChannelId     =  '19:c2560595243948b1b57291798d19dc7f@thread.skype', aChannelName  = "General"                       

       var  aThreadId      =  'AAQkADY1MDA0MGIxLTkwZDktNGE1MS1hNzM4LTQzZTNhMWNkNzEzMQAQAEEH3NGLv0VPr_mKCxbut-o='; 
       var  aThreadName    =  'review specs for File upload'; 

       var  aConversId     =  'AAQkADY1MDA0MGIxLTkwZDktNGE1MS1hNzM4LTQzZTNhMWNkNzEzMQAQAEEH3NGLv0VPr_mKCxbut-o='; 
       var  aConversName   =  'review specs for File upload'; 



       var  aTeamsAPI      =   putIDs( pAPIs.Teams.API,     { userId:    aUserId    } )   
       var  aTeamsName     =   putIDs( pAPIs.Teams.Name,    { userName:  aUserName  } )   

       var  aChannelsAPI   =   putIDs( pAPIs.Channels.API,  { teamId:    aTeamId    } )   
       var  aChannelsName  =   putIDs( pAPIs.Channels.Name, { teamName:  aTeamName  } )   

       var  aMessagesAPI   =   putIDs( pAPIs.Messages.API,  { teamId:    aTeamId  , channelId:        aChannelId   } )   
       var  aMessagesName  =   putIDs( pAPIs.Messages.Name, { teamName:  aTeamName, channelName:      aChannelName } )   

       var  aThreads1API   =   putIDs( pAPIs.Threads1.API,  { groupId:   aTeamId    } )   
       var  aThreads1Name  =   putIDs( pAPIs.Threads1.Name, { groupName: aTeamName  } )    

       var  aThreads2API   =   putIDs( pAPIs.Threads1.API,  { groupId:   aTeamId,   convserationId:   aConversId   } )   
       var  aThreads2Name  =   putIDs( pAPIs.Threads1.Name, { groupName: aTeamName, conversationName: aConversName } )    
  
       var  aConversAPI    =   putIDs( pAPIs.Convers.API,   { groupId:   aTeamId    } )   
       var  aConversName   =   putIDs( pAPIs.Convers.Name,  { groupName: aTeamName  } )    

       var  aPosts1API     =   putIDs( pAPIs.Posts2.API,    { groupId:   aTeamId,   threadId:         aThreadId         } )   
       var  aPosts1Name    =   putIDs( pAPIs.Posts2.Name,   { groupName: aTeamName, threadName:       aThreadName       } )    

       var  aPosts2API     =   putIDs( pAPIs.Posts2.API,    { groupId:   aTeamId,   convserationId:   aConversId,   threadId:   aThreadId   } )   
       var  aPosts2Name    =   putIDs( pAPIs.Posts2.Name,   { groupName: aTeamName, conversationName: aConversName, threadName: aThreadName } )    




//    ----  --------------------------------------------------------

//     var  pTeams    =  await  shoAPI(  aTeamsAPI,     aScopes1, fmtTeams3, '', aTeamsName )

//     var  pThreads  =  await  shoAPI( '/groups/{id}/threads',       aScopes1, fmtThreads3, aGroupId, aTeamName )
//     var  pThreads  =  await  shoAPI( '/groups/{id}/conversations', aScopes1, fmtThreads3, aGroupId, aTeamName )
//     var  pThreads  =  await  shoAPI( `/groups/{id}/conversations/${aConvId}/threads`, aScopes1, fmtThreads3, aConvId, aConvTeamName, 'beta' )
//     var  pThreads  =  await  shoAPI( `/conversationThreads`, aScopes1, fmtThreads3, aGroupId, aTeamName )

//     var  pChannels  =  await  shoAPI(  aChannelsAPI, aScopes1, fmtChannels3,'', aChannelsName, 'beta' ); if (pChannels.error) {  return }
//     var  pMessages  =  await  shoAPI(  aMessagesAPI, aScopes1, fmtMessages3,'', aMessagesName, 'beta' ); if (pMessages.error) {  return }
       var  pConvers   =  await  shoAPI(  aConversAPI,  aScopes1, fmtThreads3, '', aConversName,  'beta' ); if (pConvers.error ) {  return }
//     var  pThreads1  =  await  shoAPI(  aThreads1API, aScopes1, fmtThreads3, '', aThreads1Name, 'beta' ); if (pThreads1.error) {  return }
//     var  pThreads2  =  await  shoAPI(  aThreads2API, aScopes1, fmtThreads3, '', aThreads2Name, 'beta' ); if (pThreads2.error) {  return }
//     var  pPosts1    =  await  shoAPI(  aPosts1API,   aScopes1, fmtPosts3,   '', aPosts1Name,   'beta' ); if (pPosts1.error  ) {  return }
//     var  pPosts2    =  await  shoAPI(  aPosts2API,   aScopes1, fmtPosts3,   '', aPosts2Name,   'beta' ); if (pPosts2.error  ) {  return }

      }; // eof getThreads1
//    --------------------------------------------------------------------

  function  putIDs( aStr, pNames ) { 
            Object.keys(  pNames ).map( aName => aStr = aStr.replace( `{${aName}}`, pNames[ aName ] ) ) 
     return aStr       
            }

  function  fmtThreads1( mThreads, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'asData'

       var  mRecs =    mThreads.map( ( pRec, i ) => {
       var  pFlds = { "Subject"     :           pRec.topic
                    , "UpdatedOn"   : fmtDate(  pRec.lastDeliveredDateTime ) + " "
                    , "Preview"     :           pRec.preview
                    , "Emails"      :           pRec.uniqueSenders.join()
                    , "Id"          :           pRec.id
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter(   selectChannel1 )  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Threads" )
    return  mData
         }; // eof fmtThreads1
// -------  ------------------------------------------------------------------

  function  fmtThreads3( mThreads, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'byRows'

       var  mRecs =    mThreads.map( ( pRec, i ) => {
       var  pFlds = { "Subject"     : chop( 40, pRec.topic.replace( /\n/g, "<br>" ) )
                    , "UpdatedOn"   : fmtDate(  pRec.lastDeliveredDateTime ) + " "
                    , "Emails"      : chop( 30, pRec.uniqueSenders.join( ", " ) )
                    , "Id"          :           pRec.id
//                  , "Preview"     : chop( 70, pRec.preview.replace( /\n/g, "<br>" ) )
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter(   selectChannel1 )  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Threads" )
    return  mData
         }; // eof fmtThreads3
// -------  ------------------------------------------------------------------

    }; // eif threads1
// -------  -----------------------------------------------------------------------------------------------------------------------

       if (aEntities.match( /groups1/ )) {

                getGroups1() 

//    --------------------------------------------------------------------

 async function getGroups1( ) { 
       var  aTeams_API    =  '/users/{id}/joinedTeams'
       var  aChannels_API =  '/teams/{id}/channels'
       var  aMessages_API =  '/teams/{TeamId}/channels/{id}/messages'
       var  aThreads_API  =  '/groups/{id}/threads'
       var  aGroups_API   =  '/groups'

       var  aScopes1      =  ''
//     var  aScopes2      =  'Channel.ReadBasic.All'

//    ----  --------------------------------------------------------

       var  pGroups       =  await  shoAPI( aGroups_API, aScopes1, fmtGroups3 )
        if (pGroups.error) { return }

      }; // eof getGroups1
//    --------------------------------------------------------------------

  function  fmtGroups1( mGroups, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'asData'

       var  mRecs =    mGroups.map( ( pRec, i ) => {
       var  pFlds = { "Name"        :           pRec.displayName 
                    , "CreatedOn"   : fmtDate(  pRec.createdDateTime ) + " "
                    , "Email"       :           pRec.mail 
                    , "Id"          :           pRec.id
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter(   selectChannel1 )  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Groups" )
    return  mData
         }; // eof fmtGroups1
// -------  ------------------------------------------------------------------


  function  fmtGroups3( mGroups, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'byRows'

       var  mRecs =    mGroups.map( ( pRec, i ) => {
       var  pFlds = { "Name"        : chop( 40, pRec.displayName )
                    , "CreatedOn"   : fmtDate(  pRec.createdDateTime ) + " "
                    , "Email"       : chop( 40, pRec.mail )
                    , "Id"          :           pRec.id
                       }
    return  pFlds
            } )
            mRecs =   mRecs.filter(   sortGroups3)  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Groups" )
    return  mData
         }; // eof fmtThreads3
// -------  ------------------------------------------------------------------

//function  sortGroups3( a, b ) { try { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : -1 } catch(e) { return 1 } }
  function  sortGroups3( a, b ) { return a.Name.toLowerCase() > String(b.Name).toLowerCase() ? 1 : -1 }

    }; // eif fmtGroups3
// -------  -----------------------------------------------------------------------------------------------------------------------



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
       var  pFlds = { 'Name'       :          pRec.displayName
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
                    , 'Name'       : chop( 25, pRec.displayName )
                    , 'Id'         :           pRec.id
//                  , 'CreatedOn'  : fmtDate(  pRec.createdDateTime + '' )
//                  , 'WebUrl'     : chop( 35, pRec.webUrl )
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
                    , 'Name'       : chop( 25, pRec.displayName )
                    , 'Id'         :           pRec.id + " "
//                  , 'CreatedOn'  : fmtDate(  pRec.createdDateTime + '' )  // always null
//                  , 'WebUrl'     : chop( 35, pRec.webUrl )                // always null
                    , 'Description': chop( 70, pRec.description )
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter( where( 'selectTeam3' ) )  // must be global
            mRecs =   mRecs.filter(           selectTeam1    )  // must be global and defined
            mRecs =   mRecs.sort(             sortTeam1      )  // must be global and defined
//     var  mData =   fmtData( mRecs, aData, "Teams", "qq" )
       var  mData =   fmtData( mRecs, aData, "Teams" )
    return  mData

         }; // eof fmtTeams_3
// -------  ------------------------------------------------------------------

  function  sortTeam1(  a,b  ) { return a.Name.toLowerCase() > String(b.Name).toLowerCase() ? 1 : -1  }
  function  selectTeam1(  pTeam  ) {
//  return  true 
    return  pTeam.Name.match( /Support|Development/ )
            }
// -------  ------------------------------------------------------------------
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
// -------  ------------------------------------------------------------------
// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtChannels1( mChannels, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'asData'

       var  mRecs =    mChannels.map( ( pRec, i ) => {
       var  pFlds = { "Name"        :           pRec.displayName 
                    , "CreatedOn"   : fmtDate(  pRec.createdDateTime ) + " "
                    , "Description" :           pRec.description
                    , "WebUrl"      :           pRec.webUrl
                    , "Email"       :           pRec.email
                    , "Id"          :           pRec.id
                       }
    return  pFlds
            } )
//          mRecs =   mRecs.filter(   selectChannel1 )  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Channels" )
    return  mData
         }; // eof fmtChannels1
// -------  ------------------------------------------------------------------

  function  fmtChannels2( mChannels, asData ) {   // asData or byRows
       var  asData =   asData ? asData : 'byFlds'
    return  fmtChannels1( mChannels, asData )
            }

  function  fmtChannels3( mChannels, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'byRows'

       var  mRecs =    mChannels.map( ( pRec, i ) => {
       var  pFlds = { 'No'         : take( -3, i + 1 ) + "."
                    , "Name"       : chop( 33, pRec.displayName )
                    , "Id"         :           pRec.id + " "
                    , "Email"      : chop( 34, pRec.email )
                    , "CreatedOn"  : fmtDate(  pRec.createdDateTime ) + " "
//                  , "WebUrl"     :           pRec.webUrl
                    , "Description": chop( 20,  pRec.description )  
                       }
    return  pFlds
            } )
            mRecs =   mRecs.filter(   selectChannel3 )  // must be global and defined
            mRecs =   mRecs.sort(     sortChannel3   )  // must be global and defined
       var  mData =   fmtData( mRecs, aData, "Channels" )
    return  mData
         }; // eof fmtChannels1
// -------  ------------------------------------------------------------------

// - createdDateTime		: '2021-05-12T10:39:31.442Z'
// - description			:  null
// - displayName			: 'Cobblestone'
//   email				    :  null
// - id					    : '19:1e54a1c831f34e63b03c0127465077a0@thread.tacv2'
//   isFavoriteByDefault    :  null
//   membershipType		    : 'standard'
//   tenantId			    : 'b5d500c1-e385-42eb-ab34-db10f116a192'
// -  webUrl				: 'https://teams.microsoft.com/l/channel/19%3A1e54a1c831f34e63b03c0127465077a0%40thread.tacv2/Cobblestone?groupId=3ff9068a-be48-4aa3-8f39-4853f72abb7b&tenantId=b5d500c1-e385-42eb-ab34-db10f116a192&allowXTenantAccess=False'

  function  sortChannel3(   a, b  ) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : -1 } 
  function  selectChannel3( pTeam ) { return  pTeam.Name.match( /General|BASEC Issues|Customer Issues/ ) }
            
// -------  ------------------------------------------------------------------
// -------  -----------------------------------------------------------------------------------------------------------------------

if (1 == 2) {
     { Error: [ '** Forbidden'
          , `Missing role permissions on the request. API requires one of` 
          ,    `ChannelSettings.Read.All, Channel.ReadBasic.All, ChannelSettings.ReadWrite.All, Group.Read.All, 
                Directory.Read.All, Group.ReadWrite.All, Directory.ReadWrite.All, ChannelSettings.Read.Group, 
                ChannelSettings.Edit.Group, ChannelSettings.ReadWrite.Group.` 
          , `Roles on the request 'Notes.Read.All, User.Read.All, Notes.ReadWrite.All'.` 
          , `Resource specific consent grants on the request ''.`
             ]
       }       
    }         




























// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtNotes1( mData, asData ) { return fmtNotes_1( mData, asData ) }
  function  fmtNotes2( mData, asData ) { return fmtNotes_2( mData, asData ) }
  function  fmtNotes3( mData, asData ) { return fmtNotes_1( mData,'byRows') }

  function  fmtNotes_1( mNotebooks, asData ) {   // asData or byRows
       var  aData =    asData ? asData : 'asData'

       var  mRecs =    mNotebooks.map( pRec => {
       var  pFlds = { "Name"        : chop( 40, pRec.displayName )
                    , "UpdatedOn"   : fmtDate(  pRec.lastModifiedDateTime ) + " "
                    , "UpdatedBy"   :           pRec.lastModifiedBy.user.displayName
                       }
    return  pFlds
            } )
//     var  mData = fmtData( mRecs, aData, "Notebooks", 'qq' )  // bQQ  if bCSV == false
//     var  mData = fmtData( mRecs, aData, "Notebooks", true )  // bCSV if asData != 'asData'
       var  mData = fmtData( mRecs, aData, "Notebooks" )        // bCSV if asData != 'asData'
    return  mData
         }; // eof fmtNotes_1
// -------  ------------------------------------------------------------------

  function  fmtNotes_2( mNotebooks, asData ) { // byFlds
       var  aData =    asData ? asData : 'byFlds', bData = (aData == 'asData')

       var  mRecs =    mNotebooks.map( pRec => {
       var  mFlds = [ "Name:         " +          pRec.displayName
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
// -------  ------------------------------------------------------------------
// -------  -----------------------------------------------------------------------------------------------------------------------


































// -------  -----------------------------------------------------------------------------------------------------------------------

  function  fmtMessages1( mData ) { return fmtMessages_1( mData, 'asData' ) }
  function  fmtMessages2( mData ) { return fmtMessages_1( mData, 'byFlds' ) }
  function  fmtMessages3( mData ) { return fmtMessages_3( mData, 'byRows' ) }

  function  fmtMessages_1( mMessages, asData ) {  // asData or  byFlds

       var  mRecs =    mMessages.map( pRec => {
       var  pFlds = { "Subject"       :              pRec.subject
                    , "From"          :              pRec.from.user.displayName
                    , "CreatedOn"     : fmtDate( 19, pRec.createdDateTime        )     // 19 is the default
                    , "UpdatedOn"     : fmtDate(     pRec.lastModifiedDateTime + '' )  // if you want 'null
                    , "DeletedOn"     : fmtDate( 22, pRec.deletedDateTime      || '' ) // prevents fmtDate( 'now' )
                    , "Type"          :              pRec.messageType
                    , "BodyType"      :              pRec.body.contentType
                    , "BodyLength"    :              pRec.body.content.length
                    , "WebUrl"        :              pRec.webUrl
                    , "AttachmentCnt" :              pRec.attachments.length
                       }
                   if (pRec.attachments.length > 0) {     
//          pFlds.Attachments = [ ]           
                       pRec.attachments.map( ( pAttachment, i ) => {

//          pFlds.attachments.push(  
//                  { "Name" : pAttachment.name 
//                  , "Url"  : pAttachment.contentUrl 
//                  , "Id"   : pAttachment.id 
//                     } )
            pFlds[    `Attachmnent_${i+1}.Name` ] =  pAttachment.name 
            pFlds[    `Attachmnent_${i+1}.Url`  ] =  pAttachment.url
            pFlds[    `Attachmnent_${i+1}.Id`   ] =  pAttachment.id

//          mFlds.push( `AtchFile[${i+1}].Name: ` +  pAttachment.name )
//          mFlds.push( `AtchFile[${i+1}].Url : ` +  pAttachment.contentUrl )
//          mFlds.push( `AtchFile[${i+1}].Id  : ` +  pAttachment.id )

                       } )
                   }

    return  bData ? pFlds : pFlds.join( "\n  " ) // Should be pFlds
                  } )
        if (bData) { return  mRecs } else {
       var  aLine = "\n".padEnd( 53, "-" ) + "\n  "
//     var  aRecs = aLine + mRecs.join( aLine ) + aRecs + aLine.replace( /[\n ]+$/, "" )
       var  aRecs = aLine + mRecs.join( aLine ) +         aLine.replace( /[\n ]+$/, "" )
            aRecs = "\nMessages:" + aRecs
    return  aRecs }

         }; // eof fmtMessages_1
// -------  ------------------------------------------------------------------


  function  fmtMessages_3(  mMessages ) {  // byRows

       var  mRecs =    mMessages.map( ( pRec , i ) => {
       var  aUser =               pRec.from ? pRec.from.user.displayName : "N/A"
       var  mFlds = [  fmtNum( 3, i + 1 ) + ". "
                    ,  chop(  30, aUser )
                    ,  fmtDate(   pRec.lastModifiedDateTime ) + " "
//                  ,  take(  18, pRec.messageType )
                    ,             pRec.etag
                    ,             pRec.id
                    ,  fmtNum( 6, pRec.body.content ? pRec.body.content.length : 0 )
                    ,  fmtNum( 3, pRec.attachments  ? pRec.attachments.length  : 0 ) + " "
                    ,  pRec.subject
                       ]
    return  mFlds.join( " " )
                  } )
       var  aLine = "\n  "
       var  aRecs =  aLine + mRecs.join( aLine )
            aRecs = "\nMessages:" + aRecs
    return  aRecs

         }; // eof fmtMessages_2
// -------  -----------------------------------------------------------------------------------------------------------------------





























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

