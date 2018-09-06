
(function () {
    "use strict";
    $('#comment').click(comment);

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            $('#button1').click(hightlightLongestWord);

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(hightlightLongestWord);
                return;
            }

            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");
            
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            //$('#comment').click(comment);

        });
    };
    function comment(ooXml) {
        $.ajax({
            type: "POST",
            url: "/api/comments/convertOOXmlToComments",
            // The key needs to match your method's input parameter (case-sensitive).
            data: JSON.stringify(ooXml),
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (data) {
                $("#jsGrid").jsGrid({

                    height: "500px",
                    width: "100%",
                    filtering: true,
                    sorting: true,
                    paging: true,
                    autoload: true,
                    pageSize: 10,
                    pageButtonCount: 5,

                    data:data,
                    fields: [
                        { name: "Id", type: "number", width: 150, validate: "required" },
                        { name: "CommentedText", type: "text", width: 200 },
                        { name: "Date", type: "date", width: 100 },
                        { name: "Author", type: "text", width: 200 },
                        { name: "Text", type: "text", width: 200 },
                    ]
                });
             },
            failure: function (errMsg) {
                showNotification("result",data);
            }
        });
    }

    function hightlightLongestWord() {

        Word.run(function (context) {
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
                    comment(currentOOXML);
                });
        })
        .catch(errorHandler);
    } 

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
