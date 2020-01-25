import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import {getPreviousMonday} from "./dateUtils";

export interface COL {
    mon: string,
    tue: string,
    wed: string,
    thu: string,
    fri: string,
    assign: string,
    base: string
}

export interface ROW {
    day: number,
    base: number,
}

export function getColList(): COL {
    return {
        mon: 'B',
        tue: 'C',
        wed: 'D',
        thu: 'E',
        fri: 'F',
        assign: 'G',
        base: 'A',
    }
}

export function getRowList(): ROW {
    return {
        day: 1,
        base: 2,
    }
}

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

export function getLastBottomRow(sheet: GoogleAppsScript.Spreadsheet.Sheet, cell: string): number {
    const lastRow: number = sheet.getRange(cell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    return lastRow - 1;
}

export function initSheet(sheet: Sheet, startDate: Date, defaultHour: number, bottomRow: number): void {
    const col: COL = getColList();
    const row: ROW = getRowList();
    sheet.getRange(`${col.mon}${row.day}`).setValue(startDate);
    for (let i = row.base; i <= bottomRow; i++) {
        sheet.getRange(`${col.mon}${i}`).setValue(defaultHour);
        sheet.getRange(`${col.tue}${i}`).setValue(defaultHour);
        sheet.getRange(`${col.wed}${i}`).setValue(defaultHour);
        sheet.getRange(`${col.thu}${i}`).setValue(defaultHour);
        sheet.getRange(`${col.fri}${i}`).setValue(defaultHour);
        sheet.getRange(`${col.assign}${i}`).setValue(0);
    }
}