const xl = require('excel4node');
require('dotenv').config();
const fetch = require('node-fetch');

// Remplacez les variables ci-dessous par vos propres clés et ID de table
const apiKey = process.env.API_KEY;
const baseId = process.env.BASE_ID;
const tableName = process.env.TABLE_ID;

const url = `https://api.airtable.com/v0/${baseId}/${tableName}`;

async function fetchData() {
	try {
		const response = await fetch(url, {
			headers: {
				Authorization: `Bearer ${apiKey}`,
				'Content-Type': 'application/json',
			},
		});
		const data = await response.json();
		return data.records;
	} catch (error) {
		console.log(error);
	}
}

(async () => {
	console.log('Here we start');
	const data = await fetchData();
	console.log(data);

	// On crée un objet pour stocker les données filtrées par destinataire
	const filteredData = {};

	// On boucle sur les données pour filtrer par destinataire
	data.forEach(record => {
		const recipient = record.fields['Contact'];
		if (!filteredData[recipient]) {
			filteredData[recipient] = [];
		}
		filteredData[recipient].push(record.fields);
	});

	// On boucle sur les données filtrées pour créer un fichier Excel par destinataire
	Object.keys(filteredData).forEach(recipient => {
		const wb = new xl.Workbook();
		const ws = wb.addWorksheet('Worksheet Name');

		const headingColumnNames = [
			'GA Property',
			'Type',
			'Expires',
			'Website',
			'Owner',
			'IT HQ Proposition',
			'Priority Level (HQ)',
			'GCMS',
			'G24',
		];

		let headingColumnIndex = 1;
		headingColumnNames.forEach(heading => {
			ws.cell(1, headingColumnIndex++).string(heading);
		});

		// On boucle sur les données filtrées pour les ajouter au fichier Excel correspondant
		let rowIndex = 2;
		filteredData[recipient].forEach(row => {
			let columnIndex = 1;
			Object.keys(row).forEach(columnName => {
				ws.cell(rowIndex, columnIndex++).string(row[columnName]);
			});
			rowIndex++;
		});

		wb.write(`./output/${recipient.replace('/', '_')}.xlsx`);
	});
})();
