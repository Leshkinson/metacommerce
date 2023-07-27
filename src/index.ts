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
//1. Создать с помощью nodeJS отчет в Excel
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
//1. Создать с помощью nodeJS отчет в Google Sheets
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
//3. Создать с помощью nodeJS отчет в Google Sheets с помощью Google Apps Script
//Скрипт
function createReportDataFromAPI() {
    // Забираем данные из API
    const response = UrlFetchApp.fetch("https://api.publicapis.org/entries");
    const data = JSON.parse(response.getContentText());

    // Исключаем объекты с HTTPS: false
    const filteredDataFromApi = data.entries
        .filter(entry => entry.HTTPS) // if entry.HTTPS === true
        .sort((entryA, entryB) => entryA.API.localeCompare(entryB.API));

    // Получаем заголовки из API
    //const arrayHeaders = [...new Set(data.flatMap((item) => Object.keys(item)))];
    const arrayHeaders = ["API", "Description", "Auth", "HTTPS", "Cors", "Link", "Category"];
    // Получаем активный лист в Google Sheets
    const sheet = SpreadsheetApp.getActiveSheet();

    // Записываем заголовки
    sheet.appendRow(arrayHeaders);

    // Записываем данные из API
    filteredDataFromApi.forEach((entry) => {
        sheet.appendRow([entry.API, entry.Description, entry.Auth, entry.HTTPS, entry.Cors, { text: entry.Link, hyperlink: entry.Link }, entry.Category]);
    });
}

const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY, SHEET_ID } = process.env;
const createReport = new CreateReport(SHEET_ID as string);

createReport.createSheet("https://api.publicapis.org/entries", "Metacommerce", GOOGLE_SERVICE_ACCOUNT_EMAIL as string, GOOGLE_PRIVATE_KEY as string).then(() => {
    console.log("Successful")
});
createReport.createReportByExcel("https://api.publicapis.org/entries", "Metacommerce", "ccccc").then(() => {
    console.log("Successful")
});


