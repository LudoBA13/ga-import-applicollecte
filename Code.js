function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Importer')
		.addItem('Importer AssoConnect', 'importAssoConnect')
		.addItem('Importer AppliCollecte', 'importAppliCollecte')
		.addToUi();
}

function importAssoConnect()
{
	showImportDialog('AssoConnect');
}

function importAppliCollecte()
{
	showImportDialog('AppliCollecte');
}

function showImportDialog(source)
{
	const htmlOutput = HtmlService.createTemplateFromFile('ImportDialog');
	htmlOutput.source = source;

	const html = htmlOutput.evaluate()
		.setWidth(400)
		.setHeight(200);

	SpreadsheetApp.getUi().showModalDialog(html, 'Importation ' + source);
}

/**
 * Validates the imported sheet.
 * For now, it checks if the sheet is not empty.
 * This can be expanded with specific header checks for each source.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to validate.
 * @param {string} source The source name ('AssoConnect' or 'AppliCollecte').
 * @returns {boolean}
 */
function isValidFile(sheet, source)
{
	const lastRow = sheet.getLastRow();
	const lastColumn = sheet.getLastColumn();

	// Basic validation: sheet must have data
	if (lastRow < 1 || lastColumn < 1)
	{
		return false;
	}

	// Add more specific validation here if needed
	// e.g., check for mandatory headers

	return true;
}

/**
 * Processes the uploaded file.
 *
 * @param {string} base64Data
 * @param {string} fileName
 * @param {string} source
 */
function processUpload(base64Data, fileName, source)
{
	try
	{
		const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.MICROSOFT_EXCEL, fileName);

		// 1. Upload to Drive and convert to Google Sheets
		const resource = {
			title: fileName,
			mimeType: MimeType.GOOGLE_SHEETS
		};
		const tempFile = Drive.Files.insert(resource, blob, {convert: true});
		const tempId = tempFile.id;

		const tempSpreadsheet = SpreadsheetApp.openById(tempId);
		const tempSheet = tempSpreadsheet.getSheets()[0];

		// 2. Validate
		if (!isValidFile(tempSheet, source))
		{
			Drive.Files.remove(tempId);
			throw new Error('Le fichier n\'est pas valide pour l\'importation ' + source + '.');
		}

		// 3. Replace content in target sheet
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		let targetSheet = ss.getSheetByName(source);

		if (!targetSheet)
		{
			targetSheet = ss.insertSheet(source);
		}

		targetSheet.clear();

		const data = tempSheet.getDataRange().getValues();
		if (data.length > 0)
		{
			targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
		}

		// 4. Cleanup
		Drive.Files.remove(tempId);

		return 'L\'importation de ' + source + ' a été effectuée avec succès.';
	}
	catch (e)
	{
		console.error(e);
		throw new Error('Erreur lors de l\'importation : ' + e.message);
	}
}
