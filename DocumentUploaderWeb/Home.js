
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            updateStatus("-- Kiadás előtt töröld a státusz jelzést --");

            // Add a click event handler for the highlight button.
            $('#upload-button').click(uploadFile);
        });
    };

    function uploadFile() {

        try {
            Office.context.document.getFileAsync("compressed",
                { sliceSize: 100000 },
                function (result) {

                    if (result.status == Office.AsyncResultStatus.Succeeded) {

                        // Get the File object from the result.
                        var myFile = result.value;
                        var state = {
                            file: myFile,
                            counter: 0,
                            sliceCount: myFile.sliceCount
                        };
                        spinnerToggle();
                        buttonToggle();
                        getSlice(state);
                    }
                    else {
                        updateStatus(result.status);
                    }
                });
        }
        catch (e) {
            errorHandler(e);
        }
    } 

    function getSlice(state) {
        state.file.getSliceAsync(state.counter, function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                updateStatus((state.counter + 1) + " datab feltöltése a " + state.sliceCount + " darabból");
                sendSlice(result.value, state);
            }
            else {
                updateStatus(result.status);
            }
        });
    }

    function spinnerToggle() {
        //Ha szükséges töltés jelzést akkor ezt aktiválni kell
        //$(".spinner").toggle();
    }

    function buttonToggle() {
        //Ha szükséges töltés jelzést akkor ezt aktiválni kell
        //$(".Button").toggle();
    }

    function sendSlice(slice, state) {
        var data = slice.data;

        // If the slice contains data, create an HTTP request.
        if (data) {
            //Dokumentum adatainak kiszedése a dokumentumból. Nem biztos hogy ez így működni fog, ki kell próbálni.
            //var settings = Office.context.document.settings;
            //settings.set("ugyvitel-id", "1");
            //settings.get("ugyvitel-id");

            // Encode the slice data, a byte array, as a Base64 string.
            // NOTE: The implementation of myEncodeBase64(input) function isn't
            // included with this example. For information about Base64 encoding with
            // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
            var fileData = $.base64.encode(data);

            // Create a new HTTP request. You need to send the request
            // to a webpage that can receive a post.
            var request = new XMLHttpRequest();

            // Create a handler function to update the status
            // when the request has been sent.
            request.onreadystatechange = function () {
                if (request.readyState == 4) {

                    updateStatus(slice.size + " byte feltöltése.");
                    state.counter++;

                    if (state.counter < state.sliceCount) {
                        getSlice(state);
                    }
                    else {
                        closeFile(state);
                        spinnerToggle();
                        buttonToggle();
                        showNotification("Feltöltés sikeres!", "A dokumentum frissítésre került a SIGNAL Ügyvitel rendszerben!");
                    }
                }
            }

            request.open("POST", "[ugyvitel-url]");
            request.setRequestHeader("Slice-Number", slice.index);
            request.send(fileData);

        }
    }

    function closeFile(state) {
        // Close the file when you're done with it.
        state.file.closeAsync(function (result) {

            // If the result returns as a success, the
            // file has been successfully closed.
            if (result.status == "succeeded") {
                updateStatus("Fájl lezárása sikeres.");
            }
            else {
                updateStatus("Fájl lezárása sikertelen.");
            }
        });
    }


     // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function updateStatus(message) {
        //$('#upload-status').append("<div>" + message + "</div>");
    }

    function errorHandler(error) {
        showNotification("Hiba történt", error);
        if (error instanceof OfficeExtension.Error) {
            //JSON.stringify(error.debugInfo)
        }
    }

 })();
