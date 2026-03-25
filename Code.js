function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Importer')
		.addItem('Importer de AssoConnect', 'importAssoConnect')
		.addItem('Importer de AppliCollecte', 'importAppliCollecte')
		.addToUi();

	const exportAssoMenu = ui.createMenu('Exporter vers AssoConnect')
		.addItem('Groupe AppliCollecte', 'exportAssoConnectAppliCollecte')
		.addItem('Groupe Associations', 'exportAssoConnectAssociations')
		.addItem('Groupe Partenariat', 'exportAssoConnectPartenariat');

	ui.createMenu('Exporter')
		.addSubMenu(exportAssoMenu)
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

function exportAssoConnectAppliCollecte()
{
	exportAssoConnect('AppliCollecte');
}

function exportAssoConnectAssociations()
{
	exportAssoConnect('Associations');
}

function exportAssoConnectPartenariat()
{
	exportAssoConnect('Partenariat');
}

function exportAssoConnect(groupName)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheetName = 'Import-' + groupName;
	const sheet = ss.getSheetByName(sheetName);

	if (!sheet)
	{
		SpreadsheetApp.getUi().alert('La feuille "' + sheetName + '" n\'existe pas.');
		return;
	}

	const fileName = 'Import_AssoConnect_' + groupName + '_' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd_HH-mm') + '.xlsx';
	const html = HtmlService.createTemplateFromFile('DownloadDialog');
	html.fileName = fileName;
	html.groupName = groupName;

	const interface = html.evaluate()
		.setWidth(400)
		.setHeight(150);

	SpreadsheetApp.getUi().showModalDialog(interface, 'Exportation AssoConnect - ' + groupName);
}

/**
 * Generates the XLSX file and returns the download URL.
 *
 * @param {string} groupName
 * @returns {string} The base64 data of the XLSX file.
 */
function getExportData(groupName)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const exportSheet = ss.getSheetByName('Import-' + groupName);

	// Create a temporary spreadsheet
	const tempSS = SpreadsheetApp.create('TempExport');
	const tempSheet = tempSS.getSheets()[0];

	// Copy data
	const data = exportSheet.getDataRange().getValues();
	tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

	const tempId = tempSS.getId();
	SpreadsheetApp.flush();

	// Fetch the file as XLSX via Drive API
	const url = 'https://www.googleapis.com/drive/v3/files/' + tempId + '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
	const token = ScriptApp.getOAuthToken();
	const response = UrlFetchApp.fetch(url, {
		headers: {
			'Authorization': 'Bearer ' + token
		}
	});

	const base64Data = Utilities.base64Encode(response.getBlob().getBytes());

	// Cleanup
	Drive.Files.remove(tempId);

	return base64Data;
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
