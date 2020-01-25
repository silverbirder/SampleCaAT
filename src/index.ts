import * as CaAT from '@silverbirder/caat'
import {copyDate, getLatestMonday, getPreviousMonday, getLatestFriday} from "./utils";
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const BASE_COL = 'A';
const BASE_ROW = 2;
const ASSIGN_COL = 'G';

interface IMemberLocation {
    col: string,
    row: number,
    name: string,
}

function main() {
    // prepare property
    const spreadSheetId: string = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
    const templateSheetName: string = PropertiesService.getScriptProperties().getProperty('TEMPLATE_SHEET_NAME');
    if (spreadSheetId === '') {
        throw Error('Not set the SPREAD_SHEET_ID');
    }
    if (templateSheetName === '') {
        throw Error('Not set the TEMPLATE_NAME');
    }

    // Prepare date range
    const previousMonday: Date = new Date(getPreviousMonday().setHours(0, 0, 0));
    const nextMonday: Date = new Date(getLatestMonday().setHours(0, 0, 0));
    const nextFriday: Date = new Date(getLatestFriday().setHours(23, 59, 59));

    // Prepare spread sheet
    const spreadSheet: Spreadsheet = SpreadsheetApp.openById(spreadSheetId);
    const templateSheet: Sheet = spreadSheet.getSheetByName(templateSheetName);
    const backupSheetName: string = Utilities.formatDate(previousMonday, 'Asia/Tokyo', 'YYYYMMDD');
    if (spreadSheet.getSheetByName(backupSheetName)) {
        throw Error(`${backupSheetName} Sheet already exists`)
    }
    spreadSheet.setActiveSheet(templateSheet);
    spreadSheet.duplicateActiveSheet().setName(backupSheetName);

    // Prepare CaAT member config
    const cutTimeRange: Array<CaAT.IRange> = [];
    for (const d = copyDate(nextMonday); d <= nextFriday; d.setDate(d.getDate() + 1)) {
        cutTimeRange.push({
            from: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0),
            to: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 13, 0, 0),
        })
    }
    const config: CaAT.IMemberConfig = {
        startDate: nextMonday,
        endDate: nextFriday,
        everyMinutes: 15,
        cutTimeRange: cutTimeRange,
        ignore: new RegExp(null),
    };

    // Prepare member info
    const lastBottomRow: number = getLastBottomRow(templateSheet, `${BASE_COL}${BASE_ROW}`);
    const members: Array<IMemberLocation> = [];
    let nowRow: number = BASE_ROW;
    templateSheet.getRange(`${BASE_COL}${BASE_ROW}:${BASE_COL}${lastBottomRow}`)
        .getValues().forEach((row: Array<string>) => {
        row.forEach((col: string) => {
            members.push({
                col: BASE_COL,
                row: nowRow,
                name: col,
            });
        });
        nowRow++;
    });

    // Fetch the schedules!
    members.forEach((member: IMemberLocation) => {
        const caatMember: CaAT.IMember = new CaAT.Member(`${member.name}@gmail.com`, config);
        const schedules: Array<CaAT.ISchedule> = caatMember.fetchSchedules();
        const totalAssignMinutes: number = schedules.reduce((totalAssignTime: number, schedule: CaAT.ISchedule): number => {
            return totalAssignTime + schedule.assignMinute;
        }, 0);
        templateSheet.getRange(`${ASSIGN_COL}${member.row}`).setFormula(`=CEILING(${totalAssignMinutes/60}, 0.25)`);
    })
}

function getLastBottomRow(sheet: GoogleAppsScript.Spreadsheet.Sheet, cell: string) {
    const lastRow: number = sheet.getRange(cell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    return lastRow - 1;
}