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


		console.log(data);
		return data;
	} catch (error) {
		console.log(error);
	}
}



(async () => {

	console.log('Here we start');

	const data = await fetchData();

	console.log(data);

	const wb = new xl.Workbook();
	const ws = wb.addWorksheet('Worksheet Name');

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
	]


	let headingColumnIndex = 1;
	headingColumnNames.forEach(heading => {
		ws.cell(1, headingColumnIndex++)
			.string(heading)
	});


	let rowIndex = 2;
	data.forEach(record => {
		let columnIndex = 1;
		Object.keys(record).forEach(columnName => {
			ws.cell(rowIndex, columnIndex++)
				.string(record[columnName])
		});
		rowIndex++;
	});

	wb.write('filename.xlsx');



})();