/** Returns the options to be used for the generated files.
	@return ExportOptionsJPEG object
*/
function getJPEGOptions() {
    var options = new ExportOptionsJPEG();
    options.antiAliasing = false;
    options.qualitySetting = 70;
    return options;
}

try {
    // app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    if (app.documents.length > 0) {
        // Get the folder to save the files into
        var rootFolder = null;
        rootFolder = Folder.selectDialog(
            "Select root folder to save JPG files.",
            "~",
        );

        if (rootFolder != null) {
            var options, i, sourceDoc, targetFile;

            // You can tune these by changing the code in the getJPGOptions() function.

            for (i = 0; i < app.documents.length; i++) {
                sourceDoc = app.documents[i]; // returns the document object
                sourceDoc.activate();

                // Get the file to save the document as pdf into

                var destFolder = Folder(
                    String(rootFolder) + "/" + sourceDoc.name.split(".")[0],
                );

                if (!destFolder.exists) destFolder.create();

                try {
                    options = this.getJPEGOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".jpg",
                        destFolder,
                    );
                    sourceDoc.exportFile(targetFile, ExportType.JPEG, options);
                } catch (error) {
                    throw error;
                }

                sourceDoc.save(sourceDoc.fullName);
            }
            alert("Documents saved in JPG format");
        }
    } else {
        throw new Error("There are no document open!");
    }
} catch (e) {
    alert(e.message, "Script Alert", true);
}

/** Returns the file to save or export the document into.
	@param docName the name of the document
	@param ext the extension the file extension to be applied
	@param destFolder the output folder
	@return File object
*/
function getTargetFile(docName, ext, destFolder) {
    var newName = "";

    // if name has no dot (and hence no extension),
    // just append the extension
    if (docName.indexOf(".") < 0) {
        newName = docName + ext;
    } else {
        var dot = docName.lastIndexOf(".");
        newName += docName.substring(0, dot);
        newName += ext;
    }

    // Create the file object to save to
    var myFile = new File(destFolder + "/" + newName);

    // Preflight access rights
    if (myFile.open("w")) {
        myFile.close();
    } else {
        throw new Error("Access is denied");
    }
    return myFile;
}
