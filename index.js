import fs from "fs";
import axios from "axios";
import Excel from "exceljs";
// const { google } = require('googleapis');
// const {authenticate} = require('@google-cloud/local-auth');
// const sheets = google.sheets('v4');
import {GoogleSpreadsheet} from 'google-spreadsheet';
import * as dotenv from 'dotenv';

dotenv.config()

// async function getDataFromApi(url) {
//     const workbook = new Excel.Workbook();
//     const worksheet = workbook.addWorksheet("Metacommerce");
//
//     return axios.get(url).then(async response => {
//         const data = response.data.entries
//             .filter(entry => entry.HTTPS === true)
//             .sort((a, b) => a.API.localeCompare(b.API));
//         const arrayHeaders = [...new Set(data.flatMap(item => Object.keys(item)))];
//         worksheet.columns = arrayHeaders.map((header) => {
//             return {header: `${header}`, key: `${header}`, width: 20}
//         });
//         arrayHeaders.map((header, index) => {
//             worksheet.getCell(1,index + 1).fill = {
//                 type: 'pattern',
//                 pattern:'solid',
//                 fgColor:{ argb:'cccccc' }
//             };
//             worksheet.getCell(1,index + 1).border = {
//                 top: {style: 'medium', color: {argb: "000000"}},
//                 left: {style: 'medium', color: {argb: "000000"}},
//                 bottom: {style: 'medium', color: {argb: "000000"}},
//                 right: {style: 'medium', color: {argb: "000000"}},
//             }
//         });
//         data.forEach(entry => {
//             worksheet.addRow([entry.API, entry.Description, entry.Auth, entry.HTTPS, entry.Cors, { text: entry.Link, hyperlink: entry.Link }, entry.Category]);
//         });
//
//         await workbook.xlsx.writeFile('export.xlsx');
//         //Google Sheets
//     })
// }
//
// await getDataFromApi("https://api.publicapis.org/entries")
//     .then(() => console.log('Excel document generated successfully!'))
//     .catch(error => console.log('Error:', error.message));
////////////////////////////////////////////////////
// const credentials = require('./credentials.json');
const createNewSpreadsheet = async () => {
    const private_key = process.env.GOOGLE_PRIVATE_KEY;
    const googlePrivateKey = private_key.replace(/\\n/g, '\n')
    const doc = new GoogleSpreadsheet('1pUVZF8COsRBrzhfK2IIX6oSnQxyurD8BWhQfTFzX5Ow');
    await doc.useServiceAccountAuth({
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: googlePrivateKey,
    })
    await doc.create({title: 'My Spreadsheet'})
    console.log(doc.spreadsheetId);
}
await createNewSpreadsheet();
// const doc = new Spreadsheet('your_sheet_id');
// const auth = await authenticate({
//     keyfilePath: 'credentials.json',
//     scopes: ['https://www.googleapis.com/auth/spreadsheets'],
// });
// google.options({auth});
// const sheet = await doc.addWorksheet({ title: 'API data', headerValues: arrayHeaders});
//
// const resource = {
//     values: [['API', 'Description', 'Auth', 'HTTPS', 'Link'], ...values],
//     range: 'Sheet1!A1:E' + (values.length + 1),
//     majorDimension: 'ROWS',
// };
// await sheets.spreadsheets.values.update({
//     spreadsheetId: 'idmytable',
//     range: resource.range,
//     valueInputOption: 'RAW',
//     resource,
// });

// async function createExcel() {
//     const workbook = new Excel.Workbook();
//     const worksheet = workbook.addWorksheet("Metacommerce");
//
//     worksheet.columns = [
//         {header: 'Id', key: 'id', width: 10},
//         {header: 'Name', key: 'name', width: 32},
//         {header: 'D.O.B.', key: 'dob', width: 15,}
//     ];
//
//     worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
//     worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});
//
// // save under export.xlsx
//     await workbook.xlsx.writeFile('export.xlsx');
// }
//
// await createExcel();

