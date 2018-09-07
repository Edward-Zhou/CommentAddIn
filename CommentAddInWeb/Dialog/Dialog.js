(function () {
    Office.initialize = function () {
        //Office is ready
        $(document).ready(function () {
            //the document is ready
            var ooXml = localStorage.getItem("ooXml");
            if (ooXml != null) {
                xmlToComments(ooXml);
            }
        });
    };

    function loadOoXml() {
        console.log("Load Comment Button is clicked");
        var ooXml = localStorage.getItem("ooXml");
        xmlToComments(ooXml);
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
