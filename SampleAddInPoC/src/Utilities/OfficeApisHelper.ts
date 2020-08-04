/*
     Interacting with the Office document
*/

/*
    Managing the dialogs.
*/

let loginDialog: Office.Dialog;

export const signInO365 = function() {

    console.log("signInO365");
    Office.context.ui.displayDialogAsync(
        "https://localhost:3000/login/login.html",
        { height: 40, width: 30, promptBeforeOpen: false, displayInIframe: false },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
            }
            else {
                loginDialog = result.value;
                loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
                loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
            }
        }
    );

    const processLoginMessage = function(arg: { message: string, type: string }) {

        let messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.status === 'success') {

            // We now have a valid access token.
            loginDialog.close();
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
        }
    };

    const processLoginDialogEvent = function(arg:any){
        processDialogEvent(arg);
    };

};

const processDialogEvent = function(arg: { error: number, type: string }) {

    switch (arg.error) {
        case 12002:
            break;
        case 12003:
            break;
        case 12006:
            break;
        default:
            break;
    }
};

