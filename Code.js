// Génération de l'interface de COM ! NE PAS TOUCHER !
function doGet() {
	return HtmlService.createTemplateFromFile('index')
		.evaluate()
		.setFaviconUrl('https://www.academieduclimat.paris/app/themes/academie-du-climat/src/img/favicons/favicon-32x32.png')
		.setTitle('Calendrier Académie du Climat')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDataFromSpreadsheet(range, sheetChosen = 'base', filterEmpty = true) {
	try {
		let sheetIDToUse = sheetID;
		var values = Sheets.Spreadsheets.Values.batchGet(sheetIDToUse, { ranges: [range] });
		if (values.valueRanges[0].values) {
			return filterEmpty ? values.valueRanges[0].values.filter((row) => row != '') : values.valueRanges[0].values;
		} else {
			return [];
		}
	} catch (error) {
		Logger.log('Error: ' + error);
		console.error(error.message);
		return error.message;
	}
}

function convertAbregedStringToDate(dateString, lineInTable = null) {
	const monthsMap = {
		'janv.': 0,
		'févr.': 1,
		mars: 2,
		'avr.': 3,
		mai: 4,
		juin: 5,
		'juil.': 6,
		août: 7,
		'sept.': 8,
		'oct.': 9,
		'nov.': 10,
		'déc.': 11,
		janvier: 0,
		février: 1,
		avril: 3,
		juillet: 6,
		septembre: 8,
		octobre: 9,
		novembre: 10,
		décembre: 11,
	};

	try {
		let dateSplit = dateString.split(' ');
		let day = parseInt(dateSplit[1]);
		let month = monthsMap[dateSplit[2]];
		let year = parseInt(dateSplit[3]) + 2000;
		// Création de la date en spécifiant l'heure à midi pour éviter les problèmes de fuseau horaire
		return new Date(Date.UTC(year, month, day, 12, 0, 0, 0));
	} catch (error) {
		Logger.log('Erreur à la ligne ' + lineInTable);
		Logger.log(dateString);
		Logger.log('Erreur : ' + error);
		return new Error('Erreur' + error);
	}
}
