// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function () {
        //Office is ready
        $(document).ready(function () {
            //the document is ready
        });
    };
})();

var dialog;

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Domain is not trusted");
                break;
            case 12005:
                showNotification("HTTPS is required");
                break;
            case 12007:
                showNotification("A dialog is already opened.");
                break;
            default:
                showNotification(asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;
        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}


function messageHandler(arg) {
    //dialog.close();
    showNotification(arg.message);
}


function eventHandler(arg) {

    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.

    switch (arg.error) {
        case 12002:
            showNotification("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            showNotification("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            removeOoXml();

            showNotification("Dialog closed by user");
            break;
        default:
            showNotification("Undefined error in dialog window");
            break;
    }
}

function openDialog() {

    Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog/Dialog.html",
        { height: 50, width: 50 }, dialogCallback);
}

function openDialogAsIframe() {
    //IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps.
    Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog/Dialog.html",
        { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

function saveOoXml() {
    Word.run(function (context) {
        return context.sync().then(function () {
            // Create a proxy object for the document body.
            var body = context.document.body;
            // Queue a commmand to get the OOXML contents of the body.
            var bodyOOXML = body.getOoxml();
            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    var currentOOXML = "";
                    currentOOXML = bodyOOXML.value;
                    localStorage.setItem("ooXml", currentOOXML);
                    openDialogAsIframe();
                });
        });
    });
}

function removeOoXml() {
    localStorage.removeItem("ooXml");
}
