(function () {
    Office.initialize = function () {
        //Office is ready
        $(document).ready(function () {
            //the document is ready
            $("#LoadComment").click(loadOoXml);
        });
    };

    function loadOoXml() {
        console.log("Load Comment Button is clicked");
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;
            // Queue a commmand to get the OOXML contents of the body.
            var bodyOOXML = body.getOoxml();
            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    currentOOXML = bodyOOXML.value;
                    xmlToComments(currentOOXML);
                });
        })
            .catch(errorHandler);
    }
    function xmlToComments(ooXml) {
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

                    data: data,
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
                showNotification("result", data);
            }
        });
    }
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

})();
