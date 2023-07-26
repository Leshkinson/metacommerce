import axios from "axios";
import * as dotenv from 'dotenv';
import Excel, {Worksheet} from "exceljs";
import {GoogleSpreadsheet} from 'google-spreadsheet';
//import fs from "fs";
// const sheets = google.sheets('v4');
// const { google } = require('googleapis');
// const {authenticate} = require('@google-cloud/local-auth');
dotenv.config();

interface IEntry {
    API: string;
    Description: string;
    Auth: string;
    HTTPS: boolean;
    Cors: string;
    Link: string;
    Category: string;
}

class CreateReport {
    private workbook: Excel.Workbook;
    private doc: GoogleSpreadsheet;
    constructor(sheetId: string) {
        this.workbook = new Excel.Workbook();
        this.doc = new GoogleSpreadsheet(sheetId);
    }

    public async getDataFromApi(url: string) {
        const response = await axios.get(url);
        return response.data.entries
            .filter((entry: IEntry) => entry.HTTPS) // if entry.HTTPS === true
            .sort((entryA: IEntry, entryB: IEntry) => entryA.API.localeCompare(entryB.API));
    };

    public async getArrayOfHeaders(data: IEntry[]): Promise<string[]> {
        return [...new Set(data.flatMap((item: IEntry) => Object.keys(item)))]
    };

    public async createReportByExcel(url: string, title: string, fgColor: string): Promise<void> {
        const worksheet: Worksheet = this.workbook.addWorksheet(title);
        const data = await this.getDataFromApi(url);
        const arrayHeaders = await this.getArrayOfHeaders(data);
        worksheet.columns = arrayHeaders.map((header: string) => {
            return {header: `${header}`, key: `${header}`, width: 20};
        });
        arrayHeaders.map((header: string, index: number) => {
            worksheet.getCell(1, index + 1).fill = {
                type: 'pattern',
                pattern:'solid',
                fgColor:{ argb: fgColor }
            };
            worksheet.getCell(1, index + 1).border = {
                top: {style: 'medium', color: {argb: "000000"}},
                left: {style: 'medium', color: {argb: "000000"}},
                bottom: {style: 'medium', color: {argb: "000000"}},
                right: {style: 'medium', color: {argb: "000000"}},
            };
        });
        data.forEach((entry: IEntry) => {
            worksheet.addRow([entry.API, entry.Description, entry.Auth, entry.HTTPS, entry.Cors, { text: entry.Link, hyperlink: entry.Link }, entry.Category]);
        });

        await this.workbook.xlsx.writeFile('report.xlsx');
    }

    public async createSheet(url: string, title: string, GOOGLE_SERVICE_ACCOUNT_EMAIL: string, GOOGLE_PRIVATE_KEY: string): Promise<void> {
        const data = await this.getDataFromApi(url);
        const arrayHeaders = await this.getArrayOfHeaders(data);
        await this.doc.useServiceAccountAuth({
            client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL.toString(),
            private_key: GOOGLE_PRIVATE_KEY.toString()!.replace(/\\n/g, '\n'),
        });
        await this.doc.loadInfo();

        const sheet = await this.doc.addSheet({ // Create new sheet
            title: title,
            headerValues: arrayHeaders,
            tabColor: {
                red: 1.0,
                green: 0.3,
                blue: 0.4,
                alpha: 1.0
            }
        });
        const rows = data.reduce((acc: IEntry[], entry: IEntry) => {
            let item: IEntry = {
                API: entry.API,
                Description: entry.Description,
                Auth: entry.Auth,
                HTTPS: entry.HTTPS,
                Cors: entry.Cors,
                Link: entry.Link,
                Category: entry.Category
            };
            acc.push(item);

            return acc;
        }, []);

        await sheet.addRows(rows);
    }
}
const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY, SHEET_ID } = process.env;
const createReport = new CreateReport(SHEET_ID as string);
createReport.createSheet("https://api.publicapis.org/entries", "Metacommerce", GOOGLE_SERVICE_ACCOUNT_EMAIL as string, GOOGLE_PRIVATE_KEY as string).then(() => {
    console.log("Successful")
});
// createReport.createReportByExcel("https://api.publicapis.org/entries", "Metacommerce", "ccccc").then(() => {
//     console.log("Successful")
// });
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

