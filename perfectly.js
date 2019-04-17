/** Returns the options to be used for the generated files.
	@return PDFSaveOptions object
*/
function getPDFOptions() {
    var options = new PDFSaveOptions();
    return options;
}

/** Returns the options to be used for the generated files.
	@return ExportOptionsSVG object
*/
function getSVGOptions() {
    var options = new ExportOptionsSVG();
    return options;
}

/** Returns the options to be used for the generated files.
	@return ExportOptionsJPEG object
*/
function getJPEGOptions() {
    var options = new ExportOptionsJPEG();
    options.antiAliasing = false;
    options.qualitySetting = 70;
    return options;
}

/** Returns the options to be used for the generated files.
	@return ExportOptionsPhotoshop object
*/
function getPSDOptions() {
    var options = new ExportOptionsPhotoshop();
    options.saveMultipleArtboards = true;
    options.antiAliasing = false;
    options.embedICCProfile = false;
    options.writeLayers = false;
    return options;
}

/** Returns the options to be used for the generated files.
	@return ExportOptionsAutoCAD object
*/
function getDXFOptions() {
    var options = new ExportOptionsAutoCAD();
    options.version = AutoCADCompatibility.AutoCADRelease14;
    options.exportSelectedArtOnly = false;
    options.convertTextToOutlines = false;
    options.unit = AutoCADUnit.Millimeters;
    return options;
}

/** Returns the options to be used for the generated files.
	@return EPSSaveOptions object
*/
function getEPSOptions() {
    var options = new EPSSaveOptions();
    options.includeDocumentThumbnails = true;
    options.saveMultipleArtboards = false;
    return options;
}

/** Returns the options to be used for the generated files.
	@return IllustratorSaveOptions object
*/
function getAIOptions() {
    var options = new IllustratorSaveOptions();
    options.compatibility = Compatibility.ILLUSTRATOR8;
    options.flattenOutput = OutputFlattening.PRESERVEAPPEARANCE;
    return options;
}

/** Returns the options to be used for the generated files.
	@return ExportOptionsPNG24 object
*/
function getPNG24Options() {
    var options = new ExportOptionsPNG24();
    options.antiAliasing = false;
    options.transparency = false;
    options.saveAsHTML = true;
    return options;
}
try {
    // app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    if (app.documents.length > 0) {
        // Get the folder to save the files into
        var rootFolder = null;
        rootFolder = Folder.selectDialog(
            "Select root folder to save your files into. Each open window will be saved in a different folder.",
            "~",
        );

        if (rootFolder != null) {
            var options, i, sourceDoc, targetFile;

            // You can tune these by changing the code in the getPDFOptions() function.

            for (i = 0; i < app.documents.length; i++) {
                sourceDoc = app.documents[i]; // returns the document object
                sourceDoc.activate();

                // Get the file to save the document as pdf into

                var destFolder = Folder(
                    String(rootFolder) + "/" + sourceDoc.name.split(".")[0],
                );

                if (!destFolder.exists) destFolder.create();

                try {
                    options = this.getPDFOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".pdf",
                        destFolder,
                    );
                    sourceDoc.saveAs(targetFile, options);
                } catch (e) {
                    throw e;
                }

                try {
                    options = this.getEPSOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".eps",
                        destFolder,
                    );
                    sourceDoc.saveAs(targetFile, options);
                } catch (e) {
                    throw e;
                }

                try {
                    options = this.getAIOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".ai",
                        destFolder,
                    );
                    sourceDoc.saveAs(targetFile, options);
                } catch (e) {
                    throw e;
                }

                try {
                    options = this.getSVGOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".svg",
                        destFolder,
                    );
                    sourceDoc.exportFile(targetFile, ExportType.SVG, options);
                } catch (error) {
                    throw error;
                }

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

                try {
                    options = this.getPSDOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".psd",
                        destFolder,
                    );
                    sourceDoc.exportFile(
                        targetFile,
                        ExportType.PHOTOSHOP,
                        options,
                    );
                } catch (e) {
                    throw e;
                }

                try {
                    options = this.getPNG24Options();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".png",
                        destFolder,
                    );
                    sourceDoc.exportFile(targetFile, ExportType.PNG24, options);
                } catch (error) {
                    throw error;
                }

                try {
                    options = this.getDXFOptions();
                    targetFile = this.getTargetFile(
                        sourceDoc.name,
                        ".dxf",
                        destFolder,
                    );
                    sourceDoc.exportFile(
                        targetFile,
                        ExportType.AUTOCAD,
                        options,
                    );
                } catch (error) {
                    throw error;
                }

                sourceDoc.save(sourceDoc.fullName);
            }
            alert("Documents saved in 8 formats");
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
