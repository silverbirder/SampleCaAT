import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import {getPreviousMonday} from "./utils";

export interface ILocation {
    col: string,
    row: number,
    value: string,
}

export function backupSheet(): Sheet {
    // Prepare sheet
    const spreadSheet: Spreadsheet = getSpreadSheet();
    const templateSheet: Sheet = getTemplateSheet();

    const previousMonday: Date = getPreviousMonday();
    const backupSheetName: string = Utilities.formatDate(previousMonday, 'Asia/Tokyo', 'YYYYMMDD');
    if (spreadSheet.getSheetByName(backupSheetName)) {
        throw Error(`${backupSheetName} Sheet already exists`)
    }
    spreadSheet.setActiveSheet(templateSheet);
    spreadSheet.duplicateActiveSheet().setName(backupSheetName);
    return spreadSheet.getSheetByName(backupSheetName);
}

export function getTemplateSheet(): Sheet {
    const spreadSheet: Spreadsheet = getSpreadSheet();
    // Check the need properties
    const templateSheetName: string = PropertiesService.getScriptProperties().getProperty('TEMPLATE_SHEET_NAME');
    if (templateSheetName === '') {
        throw Error('Not set the TEMPLATE_NAME');
    }
    return spreadSheet.getSheetByName(templateSheetName);
}

export function getSpreadSheet(): Spreadsheet {
    const spreadSheetId: string = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
    if (spreadSheetId === '') {
        throw Error('Not set the SPREAD_SHEET_ID');
    }
    return SpreadsheetApp.openById(spreadSheetId);
}

// ex. rangeStr: A2:A5 (Column must be the same)
export function getValueByColumn(sheet: Sheet, rangeStr: string): Array<ILocation> {
    let nowRow: number = parseInt(rangeStr.slice(1, 2));
    let colName: string = rangeStr.slice(0, 1);
    const locationList: Array<ILocation> = [];
    sheet.getRange(rangeStr).getValues().forEach((row: Array<string>) => {
        row.forEach((col: string) => {
            locationList.push({
                col: colName,
                row: nowRow,
                value: col,
            });
        });
        nowRow++;
    });
    return locationList;
}