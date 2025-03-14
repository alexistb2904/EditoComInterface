// Génération de l'interface de COM ! NE PAS TOUCHER !
function doGet() {
	return HtmlService.createTemplateFromFile('index')
		.evaluate()
		.setFaviconUrl('https://www.academieduclimat.paris/app/uploads/2025/02/🧚.png')
		.setTitle('Publications Académie du Climat')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDataFromSpreadsheet(range, sheetChosen = sheetID, filterEmpty = true) {
	try {
		let sheetIDToUse = sheetChosen;
		var values = Sheets.Spreadsheets.Values.batchGet(sheetIDToUse, { ranges: [range] });
		if (values.valueRanges[0].values) {
			return filterEmpty ? values.valueRanges[0].values.filter((row) => row != '') : values.valueRanges[0].values;
		} else {
			return [];
		}
	} catch (error) {
		Logger.log(range + ' ' + sheetChosen + ' ' + filterEmpty);
		Logger.log('Error: ' + error);
		console.error(error.message);
		return error.message;
	}
}

function appendDataToSpreadsheet(range, spreadsheetId, rowsToAdd) {
	const result = Sheets.Spreadsheets.Values.append(
		{
			values: rowsToAdd,
		},
		spreadsheetId,
		range,
		{
			valueInputOption: 'RAW',
			insertDataOption: 'INSERT_ROWS',
		}
	);
	return result;
}

function convertAbregedStringToDate(dateString, lineInTable = null) {
	try {
		let dateSplit = dateString.split(' ');
		let day = parseInt(dateSplit[1]);
		let month = parseInt(dateSplit[2]) - 1;
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

/**
 * Récupère les jours fériés et les périodes scolaires pour une année donnée et une académie donnée.
 * @param {number} year - L'année pour laquelle récupérer les données.
 * @returns {Object} - Un objet contenant les jours fériés et les périodes scolaires.
 */
function getAllFerieAndScolaire(year) {
	Logger.log(Session.getActiveUser().getEmail());
	const dateGive = new Date(year, 0, 1);
	let periodeScolaire1 = dateGive.getFullYear();
	let periodeScolaire2 = dateGive.getFullYear();

	if (dateGive.getMonth() <= 7) {
		periodeScolaire1 = dateGive.getFullYear() - 1;
		periodeScolaire2 = dateGive.getFullYear();
	} else {
		periodeScolaire1 = dateGive.getFullYear();
		periodeScolaire2 = dateGive.getFullYear() + 1;
	}
	Logger.log('Année : ' + year);
	const url = `https://calendrier.api.gouv.fr/jours-feries/metropole/${year}.json`;
	const url2 = `https://data.education.gouv.fr/api/explore/v2.1/catalog/datasets/fr-en-calendrier-scolaire/records?refine=start_date%3A%22${dateGive.getFullYear()}%22&refine=location%3A%22Paris%22`;

	try {
		const responseUrl = UrlFetchApp.fetch(url).getContentText();
		const joursFeriesData = JSON.parse(responseUrl);

		const responseUrl2 = UrlFetchApp.fetch(url2).getContentText();
		const scolaireData = JSON.parse(responseUrl2);
		Logger.log(scolaireData);
		const joursFeries = Object.entries(joursFeriesData).map(([date, nom]) => ({
			date,
			nom,
		}));

		const joursScolaires = scolaireData.results.map((item) => ({
			start_date: item.start_date,
			end_date: item.end_date,
			nom: item.description,
		}));

		return { joursFeries, joursScolaires };
	} catch (error) {
		Logger.log('Erreur lors de la récupération des jours fériés : ', error.message);
		sendTextLog('Récupération des jours fériés et scolaires', error.message);
		return error.message;
	}
}

/**
 * Envoie un journal de texte avec les détails de l'erreur.
 * @param {string} title - Le titre de l'erreur.
 * @param {string} error - La description de l'erreur.
 * @param {number} [etat=0] - L'état de l'erreur (0: Pas commencé, 1: Résolu, 2: Non Résolvable, 3: Bloqué).
 */
function sendTextLog(title, error, etat = 0) {
	/*
	let log = [];
	switch (etat) {
		case 0:
			log = ['P0', title, error, Session.getActiveUser().getEmail(), new Date(), '', 'Pas commencé'];
			break;
		case 1:
			log = ['P0', title, error, Session.getActiveUser().getEmail(), new Date(), '', 'Résolu'];
			break;
		case 2:
			log = ['P0', title, error, Session.getActiveUser().getEmail(), new Date(), '', 'Non Résolvable'];
			break;
		case 3:
			log = ['P0', title, error, Session.getActiveUser().getEmail(), new Date(), '', 'Bloqué'];
			break;
	}
	// Vers le fichier de log
	// https://docs.google.com/spreadsheets/d/1YtfK_8wWLfsgkni37k9LXrvZDb4qsu_cOsiz-FnerS8
	const sheetScript = SpreadsheetApp.openById('1YtfK_8wWLfsgkni37k9LXrvZDb4qsu_cOsiz-FnerS8').getSheetByName('Log');
	sheetScript.appendRow(log);*/
}

function convertDateToAbregedStringFormat(date) {
	const daysMap = ['dim.', 'lun.', 'mar.', 'mer.', 'jeu.', 'ven.', 'sam.'];
	date = new Date(date);
	const day = date.getUTCDate(); // Get the day of the month
	const month = date.getUTCMonth() + 1; // Months are 0-indexed
	const year = date.getUTCFullYear().toString().slice(-2); // Get last two digits of the year
	const weekday = daysMap[date.getUTCDay()]; // Get the weekday

	// Format as "weekday day month year"
	return `${weekday} ${day} ${month} ${year}`;
}

function editEventInCalendar(eventData, keysEvent) {
	if (!eventData) {
		return 'Pas de données à modifier';
	}
	Logger.log(keysEvent);
	Logger.log("Modification de l'événement : " + eventData['id'] + ' ligne: ' + eventData['ligne'] + ': tableau ' + eventData['tableau']);

	const tableauUsed = eventData['tableau'];
	let ligneInTableau = eventData['ligne'] - 1; // Spreadsheet is 0-indexed
	const numeroDossier = eventData['id'];
	const allData = getDataFromSpreadsheet(tableauUsed, sheetID);
	// Check if the event's row ID matches
	const verificationEventLigne = allData[ligneInTableau][COL_ID] ? allData[ligneInTableau][COL_ID] == numeroDossier : false;

	if (!verificationEventLigne) {
		Logger.log("Ligne de l'événement non trouvée " + numeroDossier + ", recherche de l'événement...");
		ligneInTableau = allData.findIndex((row) => row[COL_ID] == numeroDossier);
		if (ligneInTableau === -1) {
			return 'Aucun événement trouvé pour le numéro de dossier ' + numeroDossier;
		} else {
			Logger.log('Événement trouvé à la ligne ' + (ligneInTableau + 1));
			eventData['ligne'] = ligneInTableau + 1;
		}
	}

	try {
		if (eventData) {
			Logger.log(eventData);
			eventData['date'] = convertDateToAbregedStringFormat(new Date(eventData['date']));
			if (eventData['heurePublication'].includes(':')) {
				eventData['heurePublication'] =
					eventData['heurePublication'].split(':')[0] + 'h' + (eventData['heurePublication'].split(':')[1] != '00' ? eventData['heurePublication'].split(':')[1] : '');
			}
			const eventDataArray = keysEvent.map((key) => eventData[key] || '');
			eventDataArray.splice(0, 2); // Retire les 2 premiers éléments
			Logger.log(eventDataArray);
			if (eventDataArray == '') {
				return 'Aucune donnée à modifier';
			}
			if (!Array.isArray(eventDataArray)) {
				return "EventDataArray n'est pas un tableau..";
			}
			//const range = `${tableauUsed}!${colonneMin}${ligneInTableau + 1}:${columnToLetter(eventDataArray[0].length)}${ligneInTableau + 1}`;
			const sheetToUpdate = SpreadsheetApp.openById(sheetID).getSheetByName(tableauUsed);
			const sheetToUpdateId = sheetToUpdate.getSheetId();
			const requests = [
				{
					updateCells: {
						range: {
							sheetId: sheetToUpdateId,
							startRowIndex: ligneInTableau,
							endRowIndex: ligneInTableau + 1,
							startColumnIndex: 0,
							endColumnIndex: eventDataArray.length,
						},
						rows: [
							{
								values: eventDataArray.map((value) => ({
									userEnteredValue: { stringValue: value },
								})),
							},
						],
						fields: 'userEnteredValue',
					},
				},
			];
			const response = Sheets.Spreadsheets.batchUpdate({ requests: requests }, sheetID);
			Logger.log(`Batch Update : ${JSON.stringify(response)}`);
			return 1;
		} else {
			return 'Aucun événement trouvé pour le numéro de dossier ' + numeroDossier;
		}
	} catch (error) {
		Logger.log("Erreur lors de la modification de l'événement : ");
		Logger.log(error.message);
		sendTextLog("Modification de l'événement", error.message);
		return error.message;
	}
}

function columnToLetter(column) {
	let letter = '';
	while (column > 0) {
		let temp = (column - 1) % 26;
		letter = String.fromCharCode(temp + 65) + letter;
		column = (column - temp - 1) / 26;
	}
	return letter;
}

function deleteEventInCalendar(eventData, keysEvent) {
	if (!eventData) {
		return 'Pas de données à supprimer';
	}
	Logger.log(keysEvent);
	Logger.log("Suppression de l'événement : " + eventData['id'] + ' ligne: ' + eventData['ligne'] + ': tableau ' + eventData['tableau']);

	const tableauUsed = eventData['tableau'];
	let ligneInTableau = eventData['ligne'] - 1; // Spreadsheet is 0-indexed
	const numeroDossier = eventData['id'];
	const allData = getDataFromSpreadsheet(tableauUsed, sheetID);
	// Check if the event's row ID matches
	const verificationEventLigne = allData[ligneInTableau][COL_ID] ? allData[ligneInTableau][COL_ID] == numeroDossier : false;

	if (!verificationEventLigne) {
		Logger.log("Ligne de l'événement non trouvée " + numeroDossier + ", recherche de l'événement...");
		ligneInTableau = allData.findIndex((row) => row[COL_ID] == numeroDossier);
		if (ligneInTableau === -1) {
			return 'Aucun événement trouvé pour le numéro de dossier ' + numeroDossier;
		} else {
			Logger.log('Événement trouvé à la ligne ' + (ligneInTableau + 1));
			eventData['ligne'] = ligneInTableau + 1;
		}
	}

	try {
		if (eventData) {
			Logger.log(eventData);
			const eventDataArray = [eventData['mois'], eventData['semaine'], 'Annulé'];
			if (eventDataArray == '') {
				return 'Aucune donnée à supprimer';
			}
			if (!Array.isArray(eventDataArray)) {
				return "EventDataArray n'est pas un tableau..";
			}
			const sheetToUpdate = SpreadsheetApp.openById(sheetID).getSheetByName(tableauUsed);
			const sheetToUpdateId = sheetToUpdate.getSheetId();
			const requests = [
				{
					updateCells: {
						range: {
							sheetId: sheetToUpdateId,
							startRowIndex: ligneInTableau,
							endRowIndex: ligneInTableau + 1,
							startColumnIndex: 0,
							endColumnIndex: eventDataArray.length,
						},
						rows: [
							{
								values: eventDataArray.map((value) => ({
									userEnteredValue: { stringValue: value },
								})),
							},
						],
						fields: 'userEnteredValue',
					},
				},
			];
			const response = Sheets.Spreadsheets.batchUpdate({ requests: requests }, sheetID);
			Logger.log(`Batch Update : ${JSON.stringify(response)}`);
			return 1;
		} else {
			return 'Aucun événement trouvé pour le numéro de dossier ' + numeroDossier;
		}
	} catch (error) {
		Logger.log("Erreur lors de la suppression de l'événement : ");
		Logger.log(error.message);
		sendTextLog("Suppression de l'événement", error.message);
		return error.message;
	}
}

function getNumeroDossier(allData) {
	try {
		let numMax = 0;
		const allNumDossier = allData.map((row) => row[COL_ID]);
		const thisDate = new Date();
		const thisYear = thisDate.getFullYear();
		const numsDossier = allNumDossier.map(function (id) {
			const thisNumDossier = id;
			if (thisNumDossier != null && thisNumDossier != undefined && thisNumDossier != '') {
				const thisNumDossierSplit = thisNumDossier.split('-');

				if (thisNumDossierSplit[0] == thisYear) {
					const thisNum = parseInt(thisNumDossierSplit[1], 10);

					if (numMax < thisNum) {
						numMax = thisNum;
					}
				}
			}
		});

		const newNum = '00' + (numMax + 1);
		const newNumDossier = thisYear + '-' + newNum.substr(-4);
		return newNumDossier;
	} catch (error) {
		sendTextLog("Ajout d'événement NuméroDossier", error.message + JSON.stringify(value));
		Logger.log(error.message);
		return 'Erreur';
	}
}

function incrementNumeroDossier(currentNumDossier) {
	const parts = currentNumDossier.split('-');
	const year = parts[0];
	const num = parseInt(parts[1], 10) + 1;
	return `${year}-${num.toString().padStart(3, '0')}`;
}

function getNombreSemaineDate(date) {
	const startOfYear = new Date(date.getFullYear(), 0, 1);
	const dayOfYear = Math.floor((date - startOfYear) / (24 * 60 * 60 * 1000)) + 1;
	return Math.ceil(dayOfYear / 7);
}

async function createEventInCalendar(eventData, eventRepeatData, tableauUsed, keysEvent) {
	if (!eventData) {
		return 'Pas de données à modifier';
	}
	Logger.log(keysEvent);
	Logger.log(eventData);

	Logger.log('Répétition : ');
	Logger.log(eventRepeatData['repeat']);
	Logger.log('Occurrence : ');
	Logger.log(eventRepeatData);
	let allData = await getDataFromSpreadsheet(tableauUsed, sheetID);
	const numeroDossier = await getNumeroDossier(allData);
	let toAdd = [];
	try {
		if (eventData['reseau'] == 'IG STORY') {
			eventData['reseau'] = 'IG';
			eventData['type'] = 'Story';
		}
		eventData['semaine'] = getNombreSemaineDate(new Date(eventData['date']));
		eventData['mois'] = new Date(eventData['date']).toLocaleString('default', { month: 'short' });
		eventData['id'] = numeroDossier;
		eventData['ligne'] = allData.length + 1;
		eventData['date'] = convertDateToAbregedStringFormat(new Date(eventData['date']));
		Logger.log("Création de l'événement : " + eventData['id'] + ' tableau ' + eventData['tableau'] + ' ligne ' + eventData['ligne']);
		if (eventData['heurePublication'].includes(':')) {
			eventData['heurePublication'] =
				eventData['heurePublication'].split(':')[0] + 'h' + (eventData['heurePublication'].split(':')[1] != '00' ? eventData['heurePublication'].split(':')[1] : '');
		}
		const eventDataArray = keysEvent.map((key) => eventData[key] || '');
		if (eventDataArray == '') {
			return 'Aucune donnée à modifier';
		}
		if (!Array.isArray(eventDataArray)) {
			return "EventDataArray n'est pas un tableau..";
		}
		eventDataArray[2] = getNombreSemaineDate(new Date(convertAbregedStringToDate(eventData['date'])));
		allData.push(eventDataArray);
		toAdd.push(eventDataArray);
	} catch (error) {
		Logger.log("Erreur lors de la modification de l'événement : ");
		Logger.log(error.message);
		sendTextLog("Modification de l'événement", error.message);
		return error.message;
	}

	// Répétition de l'événement
	// eventRepeatData['repeat'][0] : Type de répétition
	// eventRepeatData['repeat'][1] : Date de fin de répétition
	let ligneInTableauADefinir = parseInt(allData.length);
	let dernierNumeroDossier = numeroDossier;
	if (eventRepeatData['repeat'][0] != '0') {
		let eventDate = new Date(convertAbregedStringToDate(eventData['date']));
		switch (eventRepeatData['repeat'][0]) {
			case '1':
				// Répétition journalière
				Logger.log('Répétition journalière');
				while (eventDate <= new Date(eventRepeatData['repeat'][1])) {
					eventDate.setDate(eventDate.getDate() + 1);
					let newEvent = { ...eventData };
					newEvent['semaine'] = getNombreSemaineDate(eventDate);
					newEvent['date'] = convertDateToAbregedStringFormat(eventDate);
					newEvent['mois'] = eventDate.toLocaleString('default', { month: 'short' });
					ligneInTableauADefinir++;
					newEvent['ligne'] = ligneInTableauADefinir;
					dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
					newEvent['id'] = dernierNumeroDossier;
					allData.push(newEvent);
					let newEventArray = keysEvent.map((key) => newEvent[key] || '');
					toAdd.push(newEventArray);
				}
				break;
			case '2':
				// Répétition hebdomadaire
				Logger.log('Répétition hebdomadaire');
				while (eventDate <= new Date(eventRepeatData['repeat'][1])) {
					eventDate.setDate(eventDate.getDate() + 7);
					let newEvent = { ...eventData };
					newEvent['semaine'] = getNombreSemaineDate(eventDate);
					newEvent['date'] = convertDateToAbregedStringFormat(eventDate);
					newEvent['mois'] = eventDate.toLocaleString('default', { month: 'short' });
					ligneInTableauADefinir++;
					newEvent['ligne'] = ligneInTableauADefinir;
					dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
					newEvent['id'] = dernierNumeroDossier;
					allData.push(newEvent);
					let newEventArray = keysEvent.map((key) => newEvent[key] || '');
					toAdd.push(newEventArray);
				}
				break;
			case '3':
				// Répétition bihebdomadaire
				Logger.log('Répétition bihebdomadaire');
				while (eventDate <= new Date(eventRepeatData['repeat'][1])) {
					eventDate.setDate(eventDate.getDate() + 14);
					let newEvent = { ...eventData };
					newEvent['semaine'] = getNombreSemaineDate(eventDate);
					newEvent['date'] = convertDateToAbregedStringFormat(eventDate);
					newEvent['mois'] = eventDate.toLocaleString('default', { month: 'short' });
					ligneInTableauADefinir++;
					newEvent['ligne'] = ligneInTableauADefinir;
					dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
					newEvent['id'] = dernierNumeroDossier;
					allData.push(newEvent);
					let newEventArray = keysEvent.map((key) => newEvent[key] || '');
					toAdd.push(newEventArray);
				}
				break;
			case '4':
				// Répétition mensuelle
				Logger.log('Répétition mensuelle');
				while (eventDate <= new Date(eventRepeatData['repeat'][1])) {
					eventDate.setMonth(eventDate.getMonth() + 1);
					let newEvent = { ...eventData };
					newEvent['semaine'] = getNombreSemaineDate(eventDate);
					newEvent['date'] = convertDateToAbregedStringFormat(eventDate);
					newEvent['mois'] = eventDate.toLocaleString('default', { month: 'short' });
					ligneInTableauADefinir++;
					newEvent['ligne'] = ligneInTableauADefinir;
					dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
					newEvent['id'] = dernierNumeroDossier;
					allData.push(newEvent);
					let newEventArray = keysEvent.map((key) => newEvent[key] || '');
					toAdd.push(newEventArray);
				}
				break;
			case '5':
				// Répétition annuelle
				Logger.log('Répétition annuelle');
				while (eventDate <= new Date(eventRepeatData['repeat'][1])) {
					eventDate.setFullYear(eventDate.getFullYear() + 1);
					let newEvent = { ...eventData };
					newEvent['semaine'] = getNombreSemaineDate(eventDate);
					newEvent['date'] = convertDateToAbregedStringFormat(eventDate);
					newEvent['mois'] = eventDate.toLocaleString('default', { month: 'short' });
					ligneInTableauADefinir++;
					newEvent['ligne'] = ligneInTableauADefinir;
					dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
					newEvent['id'] = dernierNumeroDossier;
					allData.push(newEvent);
					let newEventArray = keysEvent.map((key) => newEvent[key] || '');
					toAdd.push(newEventArray);
				}
				break;
			default:
				console.log('Erreur de répétition, aucun paramètres données');
				break;
		}
	} else {
		Logger.log(eventRepeatData['repeat']);
		Logger.log('Aucune répétition');
	}

	// Répétition par occurrence
	let copyOfEventRepeatData = { ...eventRepeatData };
	delete copyOfEventRepeatData['repeat'];
	if (Object.keys(copyOfEventRepeatData).length > 0) {
		Logger.log('Répétition par occurrence');
		await Object.entries(copyOfEventRepeatData).forEach(async ([position, value]) => {
			Logger.log("Création de l'événement : " + ligneInTableauADefinir);
			position = copyOfEventRepeatData[position];
			let newEvent = { ...eventData };
			newEvent['semaine'] = getNombreSemaineDate(new Date(position['occurrence2']));
			newEvent['date'] = convertDateToAbregedStringFormat(new Date(position['occurrence2']));
			newEvent['mois'] = new Date(position['occurrence2']).toLocaleString('default', { month: 'short' });
			dernierNumeroDossier = incrementNumeroDossier(dernierNumeroDossier);
			newEvent['id'] = dernierNumeroDossier;
			ligneInTableauADefinir++;
			newEvent['ligne'] = ligneInTableauADefinir;
			newEvent['heurePublication'] = position['occurrence3'];
			newEvent['titre'] = position['occurrence4'];
			newEvent['type'] = position['occurrence1'] == 'IG STORY' ? 'Story' : position['occurrence5'];
			newEvent['reseau'] = position['occurrence1'] == 'IG STORY' ? 'IG' : position['occurrence1'];

			allData.push(newEvent);
			let newEventArray = keysEvent.map((key) => newEvent[key] || '');
			Logger.log(newEventArray);
			toAdd.push(newEventArray);
		});
	}
	const lines = toAdd.map((row) => row[1]);
	toAdd = toAdd.map((row) => row.slice(2));
	const result = appendDataToSpreadsheet(tableauUsed + '!A1', sheetID, toAdd);
	Logger.log(result);
	Logger.log(toAdd);
	return [1, toAdd, lines];
}

function checkIfAccess() {
	const email = Session.getActiveUser().getEmail();
	Logger.log(email);
	const AllGroup = GroupsApp.getGroups();
	for (let i = 0; i < AllGroup.length; i++) {
		if (AllGroup[i].getEmail() == groupGoogleMail) {
			Logger.log('Accès autorisé');
			return true;
		}
	}
	Logger.log('Accès non autorisé');
	return false;
}
