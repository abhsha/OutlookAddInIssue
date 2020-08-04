/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


(function(){
  
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(){
    console.log("before messageParent");
    Office.context.ui.messageParent(JSON.stringify({ status: 'success', result: "", loggedInUser : "", loggedInUserEmail : ""}));
    console.log("after messageParent");
  };

})();
