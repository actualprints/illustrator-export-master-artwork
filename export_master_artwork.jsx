/*
 * Export Master Artwork - Illustrator ExtendScript
 *
 * This script exports each artwork layer as a separate PNG at the correct size with bleed.
 * Layers should be named by card thickness (e.g., "35", "55", "75", "100", "130", "180").
 * Cut contour lines are automatically hidden during export.
 *
 * SETUP:
 * 1. Open your master artwork file in Illustrator
 * 2. Ensure each artwork is on a separate layer named by card thickness
 * 3. Run this script via File > Scripts > Other Script...
 *
 * OUTPUT:
 * Creates a folder with PNGs named like: 79x90.png, 79x91.png, etc.
 * These can be uploaded as a ZIP to the template generator.
 */

#target illustrator

// Configuration
var CONFIG = {
    exportPPI: 1400,

    // Mapping from layer name to template size group AND exterior size with bleed
    // Exterior dimensions in points (converted from mm: 1mm = 2.8346pts)
    //
    // Supports three naming conventions:
    // 1. Card thickness (e.g., "35", "55") - One Touch sizes
    // 2. WxH dimensions (e.g., "54x72", "76x102") - Any product type
    // 3. Friendly names (e.g., "toploader-narrow", "one-touch-35")
    sizeMapping: {
        // === ONE TOUCH sizes (by card thickness in pt) ===
        // Width: 30mm = 85.04pts, Heights vary by card thickness
        "35":  { sizeGroup: "79x90",  bleedWidth: 85.04, bleedHeight: 95.67 },   // 30mm x 33.75mm
        "55":  { sizeGroup: "79x91",  bleedWidth: 85.04, bleedHeight: 97.09 },   // 30mm x 34.25mm
        "75":  { sizeGroup: "79x94",  bleedWidth: 85.04, bleedHeight: 99.21 },   // 30mm x 35mm
        "100": { sizeGroup: "79x95",  bleedWidth: 85.04, bleedHeight: 100.63 },  // 30mm x 35.5mm
        "130": { sizeGroup: "79x98",  bleedWidth: 85.04, bleedHeight: 103.46 },  // 30mm x 36.5mm
        "180": { sizeGroup: "79x101", bleedWidth: 85.04, bleedHeight: 106.30 },  // 30mm x 37.5mm

        // === TOPLOADER sizes (by slot dimensions in pts) ===
        // Narrow: 54x72pt slot + 3pt bleed (1.5pt each side) - matches slot spacing of ~56.7pt
        "54x72":  { sizeGroup: "54x72",  bleedWidth: 57, bleedHeight: 75, templateGroup: "Top Loader" },
        "toploader-narrow": { sizeGroup: "54x72",  bleedWidth: 57, bleedHeight: 75, templateGroup: "Top Loader" },

        // Wide: 76x102pt slot + 3pt bleed (1.5pt each side)
        "76x102": { sizeGroup: "76x102", bleedWidth: 79, bleedHeight: 105, templateGroup: "Top Loader" },
        "64x79": { sizeGroup: "64x79", bleedWidth: 67, bleedHeight: 82, templateGroup: "Top Loader" },
        "toploader-wide": { sizeGroup: "76x102", bleedWidth: 79, bleedHeight: 105, templateGroup: "Top Loader" }
    },

    // Patterns to identify cut contour elements (case insensitive)
    cutContourPatterns: ["cutcontour", "cut contour", "dieline", "die line", "diecut", "die cut", "kiss cut", "kisscut", "contour", "thru-cut", "thru cut", "throughcut", "through cut", "perf", "score"]
};


/**
 * Detect product type from layer names in the document.
 * Returns: "onetouch", "toploader-narrow", "toploader-wide", or null
 */
function detectProductType(doc) {
    var hasOneTouch = false;
    var hasTopLoaderNarrow = false;
    var hasTopLoaderWide = false;

    // One Touch layer names (card thickness)
    var oneTouchLayers = ["35", "55", "75", "100", "130", "180"];
    // One Touch size patterns (79xNN)
    var oneTouchPattern = /^79x\d+$/;

    for (var i = 0; i < doc.layers.length; i++) {
        var layerName = doc.layers[i].name.replace(/^\s+|\s+$/g, '').toLowerCase();

        // Check for One Touch
        for (var j = 0; j < oneTouchLayers.length; j++) {
            if (layerName === oneTouchLayers[j]) {
                hasOneTouch = true;
                break;
            }
        }
        if (oneTouchPattern.test(layerName)) {
            hasOneTouch = true;
        }

        // Check for Toploader Narrow
        if (layerName === "54x72" || layerName === "toploader-narrow" || layerName === "toploader_narrow") {
            hasTopLoaderNarrow = true;
        }

        // Check for Toploader Wide
        if (layerName === "76x102" || layerName === "64x79" || layerName === "toploader-wide" || layerName === "toploader_wide") {
            hasTopLoaderWide = true;
        }
    }

    // Return detected type (prioritize more specific detections)
    if (hasTopLoaderNarrow && !hasTopLoaderWide && !hasOneTouch) {
        return "toploader-narrow";
    }
    if (hasTopLoaderWide && !hasTopLoaderNarrow && !hasOneTouch) {
        return "toploader-wide";
    }
    if (hasOneTouch && !hasTopLoaderNarrow && !hasTopLoaderWide) {
        return "onetouch";
    }
    // Mixed or unknown - return null (no suffix)
    return null;
}


(function() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var docName = doc.name.replace(/\.[^\.]+$/, '');

    // Base path for exports
    var basePath = "C:/Users/MasterBrader/Dropbox/STICKERs/ActualPrints";

    // Prompt for proof number
    var proofNumber = prompt("Enter proof number (e.g., 6843470155 or 6843470155-1):", "");

    if (proofNumber === null || proofNumber === "") {
        alert("Proof number is required.");
        return;
    }

    // Detect product type from layers BEFORE creating folders
    var detectedProductType = detectProductType(doc);
    var productSuffix = detectedProductType ? "-" + detectedProductType : "";

    // Extract root folder name (everything before the dash, or full number if no dash)
    var rootFolder = proofNumber.split("-")[0];

    // Create output folder at base path with ROOT folder name
    var outputFolder = new Folder(basePath + "/" + rootFolder);
    if (!outputFolder.exists) {
        outputFolder.create();
    }

    // Include product type in export folder name to prevent overwrites
    var exportFolderName = proofNumber + productSuffix + "_exports";
    var exportFolder = new Folder(outputFolder.fsName + "/" + exportFolderName);
    if (!exportFolder.exists) {
        exportFolder.create();
    }

    // List all spot colors for debugging
    var spotColorNames = [];
    for (var s = 0; s < doc.spots.length; s++) {
        spotColorNames.push(doc.spots[s].name);
    }

    var result = exportByLayers(doc, exportFolder, spotColorNames);

    // Auto-zip the folder
    if (result.exported > 0) {
        var zipPath = outputFolder.fsName + "/" + proofNumber + productSuffix + "_exports.zip";
        var zipCreated = createZip(exportFolder.fsName, zipPath);

        // Save PDF copy of the master artwork
        var pdfPath = outputFolder.fsName + "/" + proofNumber + productSuffix + "-MasterArtwork.pdf";
        var pdfSaved = savePDF(doc, pdfPath);

        var msg = "Exported " + result.exported + " artworks.\n";
        msg += "Product type: " + (detectedProductType || "auto") + "\n\n";
        if (zipCreated) {
            msg += "ZIP: " + zipPath + "\n";
        } else {
            msg += "ZIP failed - please zip manually\n";
        }
        if (pdfSaved) {
            msg += "PDF: " + pdfPath + "\n";
        } else {
            msg += "PDF save failed\n";
        }
        msg += "\nHid " + result.hiddenItems + " cut contour items.\n\n" + result.details.join("\n");

        alert(msg);
    }
})();


function savePDF(doc, pdfPath) {
    try {
        // Restore all layers visibility before saving PDF
        for (var i = 0; i < doc.layers.length; i++) {
            doc.layers[i].visible = true;
        }

        var pdfFile = new File(pdfPath);
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.compatibility = PDFCompatibility.ACROBAT5;
        pdfOptions.preserveEditability = true;
        pdfOptions.generateThumbnails = true;

        doc.saveAs(pdfFile, pdfOptions);
        return true;
    } catch (e) {
        return false;
    }
}


function createZip(folderPath, zipPath) {
    try {
        var isWindows = ($.os.indexOf("Windows") !== -1);

        // Delete existing zip if present
        var existingZip = new File(zipPath);
        if (existingZip.exists) {
            existingZip.remove();
        }

        if (isWindows) {
            // Convert paths to Windows format
            var winFolderPath = folderPath.replace(/\//g, '\\');
            var winZipPath = zipPath.replace(/\//g, '\\');

            // Create a VBScript to zip the folder (works on all Windows versions)
            var vbsFile = new File(Folder.temp.fsName + "/zip_folder.vbs");
            vbsFile.open("w");
            vbsFile.writeln('Set objShell = CreateObject("Shell.Application")');
            vbsFile.writeln('Set fso = CreateObject("Scripting.FileSystemObject")');
            vbsFile.writeln('');
            vbsFile.writeln('zipPath = "' + winZipPath + '"');
            vbsFile.writeln('folderPath = "' + winFolderPath + '"');
            vbsFile.writeln('');
            vbsFile.writeln('Set objFile = fso.CreateTextFile(zipPath, True)');
            vbsFile.writeln('objFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)');
            vbsFile.writeln('objFile.Close');
            vbsFile.writeln('');
            vbsFile.writeln('Set zipFile = objShell.NameSpace(zipPath)');
            vbsFile.writeln('Set sourceFolder = objShell.NameSpace(folderPath)');
            vbsFile.writeln('');
            vbsFile.writeln('If Not zipFile Is Nothing And Not sourceFolder Is Nothing Then');
            vbsFile.writeln('    zipFile.CopyHere sourceFolder.Items, 16');
            vbsFile.writeln('    WScript.Sleep 3000');
            vbsFile.writeln('End If');
            vbsFile.close();

            // Execute the VBScript
            vbsFile.execute();

            // Wait for zip to complete
            $.sleep(4000);

            // Clean up
            vbsFile.remove();
        } else {
            // Mac: use zip command via shell script
            var shFile = new File(Folder.temp.fsName + "/zip_script.sh");
            shFile.open("w");
            shFile.writeln('#!/bin/bash');
            shFile.writeln('cd "' + folderPath + '"');
            shFile.writeln('zip -r "' + zipPath + '" .');
            shFile.close();

            shFile.execute();
            $.sleep(2000);
            shFile.remove();
        }

        // Check if zip was created
        var zipFile = new File(zipPath);
        return zipFile.exists;
    } catch (e) {
        return false;
    }
}


function exportByLayers(doc, exportFolder, spotColorNames) {
    var exported = 0;
    var results = [];
    var hiddenItemsTotal = 0;

    // Store original state
    var originalVisibility = [];
    for (var i = 0; i < doc.layers.length; i++) {
        originalVisibility.push(doc.layers[i].visible);
    }

    // Find layers with numeric names (card thickness)
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        var layerName = layer.name.replace(/^\s+|\s+$/g, ''); // trim

        // Check if layer name is in our mapping or is a WxH format
        var mapping = CONFIG.sizeMapping[layerName];
        var outputName, bleedWidth, bleedHeight;

        if (mapping) {
            outputName = mapping.sizeGroup;
            bleedWidth = mapping.bleedWidth;
            bleedHeight = mapping.bleedHeight;
        } else if (layerName.match(/^\d+x\d+$/)) {
            // Already in WxH format - add standard 6pt bleed (3 each side)
            outputName = layerName;
            var parts = layerName.split('x');
            bleedWidth = parseFloat(parts[0]) + 6;
            bleedHeight = parseFloat(parts[1]) + 6;
        } else if (layerName.match(/^\d+$/)) {
            // Unknown numeric layer - skip with warning
            results.push(layerName + ": SKIPPED (not in size mapping)");
            continue;
        } else {
            continue;
        }

        // Hide all layers except this one
        for (var j = 0; j < doc.layers.length; j++) {
            doc.layers[j].visible = (j === i);
        }

        // Hide ALL cut contour items recursively on this layer
        var hiddenItems = [];
        hideAllCutContourItems(layer, hiddenItems);
        hiddenItemsTotal += hiddenItems.length;

        // Get the CENTER of the layer content
        var contentBounds = getLayerBounds(layer);
        if (!contentBounds) {
            restoreItems(hiddenItems);
            results.push(layerName + ": SKIPPED (empty layer)");
            continue;
        }

        // Calculate center of the content
        var centerX = (contentBounds[0] + contentBounds[2]) / 2;
        var centerY = (contentBounds[1] + contentBounds[3]) / 2;

        // Create artboard at the BLEED size, centered on the content
        var artboardRect = [
            centerX - bleedWidth / 2,   // left
            centerY + bleedHeight / 2,  // top (Illustrator Y is inverted)
            centerX + bleedWidth / 2,   // right
            centerY - bleedHeight / 2   // bottom
        ];

        // Create a temporary artboard
        var originalArtboardIndex = doc.artboards.getActiveArtboardIndex();
        doc.artboards.add(artboardRect);
        var tempArtboardIndex = doc.artboards.length - 1;
        doc.artboards.setActiveArtboardIndex(tempArtboardIndex);

        // Export
        var filename = outputName + ".png";
        var exportFile = new File(exportFolder.fsName + "/" + filename);

        var options = new ExportOptionsPNG24();
        options.antiAliasing = true;
        options.artBoardClipping = true;
        options.transparency = false;

        var scale = (CONFIG.exportPPI / 72) * 100;
        options.horizontalScale = scale;
        options.verticalScale = scale;

        doc.exportFile(exportFile, ExportType.PNG24, options);

        // Remove temporary artboard
        doc.artboards.remove(tempArtboardIndex);
        doc.artboards.setActiveArtboardIndex(originalArtboardIndex);

        // Restore hidden cut contour items
        restoreItems(hiddenItems);

        exported++;
        results.push(layerName + " -> " + outputName + ".png (" + bleedWidth + "x" + bleedHeight + "pt with bleed)");
    }

    // Restore layer visibility
    for (var i = 0; i < doc.layers.length; i++) {
        doc.layers[i].visible = originalVisibility[i];
    }

    if (exported === 0) {
        alert("No layers found with valid size names.\n\nPlease name your layers with card thickness (e.g., '35', '55', '75').\n\nSpot colors found: " + spotColorNames.join(", "));
    }

    return {
        exported: exported,
        hiddenItems: hiddenItemsTotal,
        details: results
    };
}


function hideAllCutContourItems(container, hiddenItems) {
    var items = container.pageItems;

    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        if (item.hidden) continue;

        if (shouldHideItem(item)) {
            item.hidden = true;
            hiddenItems.push(item);
            continue;
        }

        if (item.typename === "GroupItem") {
            hideAllCutContourItemsInGroup(item, hiddenItems);
        }
    }
}


function hideAllCutContourItemsInGroup(group, hiddenItems) {
    for (var i = 0; i < group.pageItems.length; i++) {
        var item = group.pageItems[i];

        if (item.hidden) continue;

        if (shouldHideItem(item)) {
            item.hidden = true;
            hiddenItems.push(item);
            continue;
        }

        if (item.typename === "GroupItem") {
            hideAllCutContourItemsInGroup(item, hiddenItems);
        }
    }
}


function shouldHideItem(item) {
    if (item.name && isCutContourName(item.name)) {
        return true;
    }

    try {
        if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
            if (item.stroked && item.strokeColor) {
                var strokeColor = item.strokeColor;
                if (strokeColor.typename === "SpotColor") {
                    if (isCutContourName(strokeColor.spot.name)) {
                        return true;
                    }
                }
            }
            if (item.filled && item.fillColor) {
                var fillColor = item.fillColor;
                if (fillColor.typename === "SpotColor") {
                    if (isCutContourName(fillColor.spot.name)) {
                        return true;
                    }
                }
            }
        }
    } catch (e) {}

    if (item.typename === "GroupItem") {
        if (item.name && isCutContourName(item.name)) {
            return true;
        }
    }

    return false;
}


function isCutContourName(name) {
    if (!name) return false;
    var lowerName = name.toLowerCase();
    for (var i = 0; i < CONFIG.cutContourPatterns.length; i++) {
        if (lowerName.indexOf(CONFIG.cutContourPatterns[i]) >= 0) {
            return true;
        }
    }
    return false;
}


function restoreItems(items) {
    for (var i = 0; i < items.length; i++) {
        items[i].hidden = false;
    }
}


function getLayerBounds(layer) {
    if (layer.pageItems.length === 0) {
        return null;
    }

    var boundsObj = {left: Infinity, top: -Infinity, right: -Infinity, bottom: Infinity};
    collectBoundsRecursive(layer.pageItems, boundsObj);

    if (boundsObj.left === Infinity) {
        return null;
    }

    return [boundsObj.left, boundsObj.top, boundsObj.right, boundsObj.bottom];
}


function collectBoundsRecursive(items, boundsObj) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.hidden) continue;

        try {
            var b = item.geometricBounds;
            if (b[0] < boundsObj.left) boundsObj.left = b[0];
            if (b[1] > boundsObj.top) boundsObj.top = b[1];
            if (b[2] > boundsObj.right) boundsObj.right = b[2];
            if (b[3] < boundsObj.bottom) boundsObj.bottom = b[3];
        } catch (e) {}
    }
}
