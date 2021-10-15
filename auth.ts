// // Create the main MSAL instance
// // configuration parameters are located in config.js
//
// const msalClient = new msal.PublicClientApplication(msalConfig);
//
// async function signIn() {
//     // Login
//     try {
//         // Use MSAL to login
//         const authResult = await msalClient.loginPopup(msalRequest);
//         console.log('id_token acquired at: ' + new Date().toString());
//
//         // TEMPORARY
//         updatePage(Views.error, {
//             message: 'Login successful',
//             debug: `Token: ${authResult.accessToken}`
//         });
//     } catch (error) {
//         console.log(error);
//         updatePage(Views.error, {
//             message: 'Error logging in',
//             debug: error
//         });
//     }
// }
//
//
// function signOut() {
//     sessionStorage.removeItem('graphUser');
//     msalClient.logout();
// }
