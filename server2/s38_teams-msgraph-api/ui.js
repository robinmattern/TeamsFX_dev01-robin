
async function displayUI() {    
    
     await  signIn();

     const  user        =  await getUser();  // Display info from user profile
       var  userName    =  document.getElementById( 'userName' );
            userName.innerText = user.displayName;  
    
       var  signInButton=  document.getElementById( 'signin'   );   // Hide login button and initial UI
            signInButton.style = "display: none";
    
       var  content     =  document.getElementById( 'content'  );
            content.style = "display: block";
            }
