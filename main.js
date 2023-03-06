const xl = require('excel4node');
require('dotenv').config();
const fetch = require('node-fetch');
const fs = require('fs');

// Remplacez les variables ci-dessous par vos propres clés et ID de table
const apiKey = process.env.API_KEY;
const baseId = process.env.BASE_ID;
const tableName = process.env.TABLE_ID;

const url = `https://api.airtable.com/v0/${baseId}/${tableName}`;

async function fetchData(offset = null, records = []) {
	try {
		let apiUrl = url;
		if (offset) {
			apiUrl += `?offset=${offset}`;
		}
		const response = await fetch(apiUrl, {
			headers: {
				Authorization: `Bearer ${apiKey}`,
				'Content-Type': 'application/json',
			},
		});
		const data = await response.json();
		records.push(...data.records);
		if (data.offset) {
			return fetchData(data.offset, records);
		} else {
			return records;
		}
	} catch (error) {
		console.log(error);
	}
}

function loadJSON(filepath) {
	return JSON.parse(fs.readFileSync(filepath));
}

function writeJSON(data, filepath) {
	fs.writeFile(filepath, JSON.stringify(data), (err) => {
		if (err)
			throw err;
		console.log('The file has been saved!');
	});
}



const localJSON = './input/data.json';
(async () => {
	console.log('Here we start');


	const callAPI = false;

	let data = null;

	if (callAPI) {

		data = await fetchData();

		// écriture du JSON dans un fichier
		writeJSON(data, localJSON);
	}
	else {

		data = loadJSON(localJSON);
	}

	const recipients = data.map(record => record.fields['Contact']);
	const uniqueRecipients = [...new Set(recipients)].sort();

	// POUR TESTS : seulement une destinataire
	// const uniqueRecipients = [];
	// uniqueRecipients.push("John Doe");

	console.log(uniqueRecipients);


	uniqueRecipients.forEach(recipient => {

		const wb = new xl.Workbook();

		const migrateStyle = wb.createStyle({
			font: {
				color: '#086608',
			},
			fill: {
				type: 'pattern',
				patternType: 'solid',
				bgColor: '#c6efce',
				fgColor: '#c6efce',
			}
		});

		const feedbackStyle = wb.createStyle({
			font: {
				color: '#9c5700',
			},
			fill: {
				type: 'pattern',
				patternType: 'solid',
				bgColor: '#ffeb9c',
				fgColor: '#ffeb9c',
			}
		});

		const noMigrateStyle = wb.createStyle({
			font: {
				color: '#9c0006',
			},
			fill: {
				type: 'pattern',
				patternType: 'solid',
				bgColor: '#ffc7ce',
				fgColor: '#ffc7ce',
			}
		});

		const recipientData = data.filter(record => record.fields['Contact'] === recipient);

		const ws = wb.addWorksheet(recipient.replaceAll('/', '_'));

		const headingColumnNames = [
			"GA Property",
			"Type",
			"Expires",
			"Website",
			"Owner",
			"IT HQ Proposition",
			"Priority Level (HQ)",
			"GCMS",
			"G24",
		];

		let headingColumnIndex = 1;
		headingColumnNames.forEach(heading => {
			ws.cell(1, headingColumnIndex++)
				.string(heading)
		});

		let rowIndex = 2;
		recipientData.forEach(record => {
			ws.cell(rowIndex, 1).string(record.fields["GA Property"] ? record.fields["GA Property"].toString() : "");
			ws.cell(rowIndex, 2).string(record.fields.Type ? record.fields.Type.toString() : "");
			ws.cell(rowIndex, 3).string(record.fields.Expires ? record.fields.Expires.toString() : "");
			ws.cell(rowIndex, 4).string(record.fields.Website ? record.fields.Website.toString() : "");
			ws.cell(rowIndex, 5).string(record.fields["Owner "] ? record.fields["Owner "].toString() : "");
			ws.cell(rowIndex, 6).string(record.fields["IT HQ Proposition"] ? record.fields["IT HQ Proposition"].toString() : "");

			if (record.fields["IT HQ Proposition"] && record.fields["IT HQ Proposition"].toString() === "Migrate") {
				ws.cell(rowIndex, 6).style(migrateStyle);
			}
			if (record.fields["IT HQ Proposition"] && record.fields["IT HQ Proposition"].toString() === "Waiting for feedback") {
				ws.cell(rowIndex, 6).style(feedbackStyle);
			}
			if (record.fields["IT HQ Proposition"] && record.fields["IT HQ Proposition"].toString() === "Do not migrate" || record.fields["IT HQ Proposition"] && record.fields["IT HQ Proposition"].toString() === "Kill") {
				ws.cell(rowIndex, 6).style(noMigrateStyle);
			}

			ws.cell(rowIndex, 7).string(record.fields["Priority Level (HQ)"] ? record.fields["Priority Level (HQ)"].toString() : "");
			ws.cell(rowIndex, 8).string(record.fields.GCMS ? record.fields.GCMS.toString() : "");
			ws.cell(rowIndex, 9).string(record.fields.G24 ? record.fields.G24.toString() : "");
			rowIndex++;
		});



		wb.write(`./output/${recipient.replaceAll('/', '_')}.xlsx`);
	});
})();




