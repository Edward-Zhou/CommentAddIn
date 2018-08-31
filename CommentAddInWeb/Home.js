
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

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(displaySelectedText);
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

        //$.get("/api/comments", function (data) {
        //    showNotification("result",data);
        //});
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
    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;
            // Queue a commmand to get the OOXML contents of the body.
            var bodyOOXML = body.getOoxml();

            // Queue a commmand to clear the contents of the body.
            //body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync().then(function () {
                var currentOOXML = "";
                currentOOXML = bodyOOXML.value;
                $(function () {
                    $("#jsGrid").jsGrid({
                        height: "auto",
                        width: "100%",

                        sorting: true,
                        paging: false,
                        autoload: true,
                        loadData: function (filter) {
                            return $.ajax({
                                type: "POST",
                                url: "/api/comments/convertOOXmlToComments",
                                // The key needs to match your method's input parameter (case-sensitive).
                                data: JSON.stringify(currentOOXML),
                                contentType: "application/json; charset=utf-8",
                                dataType: "json",
                            });
                        },

                        fields: [
                            //{ name: "Id", type: "number", width: 150, validate: "required" },
                            //{ name: "Date", type: "date", width: 50 },
                            { name: "Author", type: "text", width: 200 },
                            //{ name: "Text", type: "text", width: 200 },
                            //{ type: "control" }
                        ]
                    });

                });

            });
        })
        .catch(errorHandler);
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


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
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
