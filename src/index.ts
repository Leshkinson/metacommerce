import fs from "fs";
import axios from "axios";
import Excel, {Workbook, Worksheet} from "exceljs";
// const { google } = require('googleapis');
// const {authenticate} = require('@google-cloud/local-auth');
// const sheets = google.sheets('v4');
import {GoogleSpreadsheet} from 'google-spreadsheet';
import * as dotenv from 'dotenv';
interface IEntry {
    API: string,
    Description: string,
    Auth: string,
    HTTPS: boolean,
    Cors: string,
    Link: string,
    Category: string
}

dotenv.config()

class CreateReport {
    private workbook: Workbook;
    constructor() {
        this.workbook = new Excel.Workbook();
    }
    static async getDataFromApi(url: string) {
        const response = await axios.get(url);
        return response.data.entries
            .filter((entry: IEntry) => entry.HTTPS === true)
            .sort((a: IEntry, b: IEntry) => a.API.localeCompare(b.API));
    }

    static async getArrayOfHeaders(data: IEntry[]) {
        return [...new Set(data.flatMap((item: IEntry) => Object.keys(item)))]
    }

    static async createReportByExcel(url: string, fgColor: string) {
        const workbook = new Excel.Workbook();
        const worksheet: Worksheet = workbook.addWorksheet("Metacommerce");
        const data = await this.getDataFromApi(url)
        const arrayHeaders = await this.getArrayOfHeaders(data);
        worksheet.columns = arrayHeaders.map((header) => {
            return {header: `${header}`, key: `${header}`, width: 20}
        });
        arrayHeaders.map((header: string, index: number) => {
            worksheet.getCell(1,index + 1).fill = {
                type: 'pattern',
                pattern:'solid',
                fgColor:{ argb: fgColor }
            };
            worksheet.getCell(1,index + 1).border = {
                top: {style: 'medium', color: {argb: "000000"}},
                left: {style: 'medium', color: {argb: "000000"}},
                bottom: {style: 'medium', color: {argb: "000000"}},
                right: {style: 'medium', color: {argb: "000000"}},
            }
        });
        data.forEach((entry: IEntry) => {
            worksheet.addRow([entry.API, entry.Description, entry.Auth, entry.HTTPS, entry.Cors, { text: entry.Link, hyperlink: entry.Link }, entry.Category]);
        });

        await workbook.xlsx.writeFile('export.xlsx');
    }

    static async createSheet(url: string) {
        const data = await this.getDataFromApi(url);
        const arrayHeaders = await this.getArrayOfHeaders(data);
        const doc = new GoogleSpreadsheet('1pUVZF8COsRBrzhfK2IIX6oSnQxyurD8BWhQfTFzX5Ow');
        const { GOOGLE_PRIVATE_KEY, GOOGLE_SERVICE_ACCOUNT_EMAIL } = process.env;
        await doc.useServiceAccountAuth({
            client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL as string,
            private_key: GOOGLE_PRIVATE_KEY!.replace(/\\n/g, '\n'),
        });
        await doc.loadInfo()

        const sheet = await doc.addSheet({ // Создаем новый лист
            title: 'Metacommerce',
            headerValues: arrayHeaders,
        });
        const rows = data.reduce(async (acc: IEntry[], entry: IEntry) => {
            let item: IEntry = { API: entry.API, Description: entry.Description, Auth: entry.Auth, HTTPS: entry.HTTPS, Cors: entry.Cors, Link: entry.Link, Category: entry.Category }
            acc.push(item)
            return acc
        }, [])
        await sheet.addRows(rows)

    }
}
 // CreateReport.createReportByExcel("https://api.publicapis.org/entries", "cccccc").then(() => {
 //     console.log("Successful")
 // })

CreateReport.createSheet("https://api.publicapis.org/entries").then(() => {
    console.log("Successful")
});
// const doc = new GoogleSpreadsheet();
// const { GOOGLE_PRIVATE_KEY, GOOGLE_SERVICE_ACCOUNT_EMAIL } = process.env;
//
// doc.useServiceAccountAuth({
//     client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL as string,
//     private_key: GOOGLE_PRIVATE_KEY!.replace(/\\n/g, '\n'),
// })
// async function getDataFromApi(url: string) {
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
// const createNewSpreadsheet = async () => {
//     const private_key = process.env.GOOGLE_PRIVATE_KEY;
//     const googlePrivateKey = private_key.replace(/\\n/g, '\n')
//     const doc = new GoogleSpreadsheet('1pUVZF8COsRBrzhfK2IIX6oSnQxyurD8BWhQfTFzX5Ow');
//     await doc.useServiceAccountAuth({
//         client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
//         private_key: googlePrivateKey,
//     })
//     await doc.create({title: 'My Spreadsheet'})
//     console.log(doc.spreadsheetId);
// }
// await createNewSpreadsheet();
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

