import * as CaAT from '@silverbirder/caat'
import {copyDate, getLatestMonday, skipDay} from "./utils";
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

function main() {
    // prepare property
    const userId: string = PropertiesService.getScriptProperties().getProperty('USER_ID');
    const spreadSheetId: string = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
    const templateSheetName: string = PropertiesService.getScriptProperties().getProperty('TEMPLATE_NAME');
    if (userId === '') {
        throw Error('Not set the USER_ID');
    }
    if (spreadSheetId === '') {
        throw Error('Not set the SPREAD_SHEET_ID');
    }
    if (templateSheetName === '') {
        throw Error('Not set the TEMPLATE_NAME');
    }

    // Prepare date range
    const monday: Date = new Date(getLatestMonday().setHours(0, 0, 0));
    const friday: Date = new Date(skipDay(monday, 4).setHours(23, 59, 59));

    // Prepare spread sheet
    const spreadSheet: Spreadsheet = SpreadsheetApp.openById(spreadSheetId);
    const templateSheet: Sheet = spreadSheet.getSheetByName(templateSheetName);
    const copyToSheetName: string = `${monday.getFullYear()}.${monday.getMonth() + 1}.${monday.getDate()}`;
    if (spreadSheet.getSheetByName(copyToSheetName)) {
        throw Error(`${copyToSheetName} Sheet already exists`)
    }
    const targetSheet: Sheet = templateSheet.copyTo(spreadSheet);
    targetSheet.setName(copyToSheetName);

    // Prepare member object
    const cutTimeRange: Array<CaAT.IRange> = [];
    for (const d = copyDate(monday); d <= friday; d.setDate(d.getDate() + 1)) {
        cutTimeRange.push({
            from: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0),
            to: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 13, 0, 0),
        })
    }
    const config: CaAT.IMemberConfig = {
        startDate: monday,
        endDate: friday,
        everyMinutes: 15,
        cutTimeRange: cutTimeRange,
        ignore: new RegExp(null),
    };
    const member: CaAT.IMember = new CaAT.Member(userId, config);

    // Fetch the schedules!
    const schedules: Array<CaAT.ISchedule> = member.fetchSchedules();
    const totalAssignMinutes: number = schedules.reduce((totalAssignTime: number, schedule: CaAT.ISchedule): number => {
        return totalAssignTime + schedule.assignMinute;
    }, 0);

    // Set schedules to spread sheet
    const cell: Array<Array<string>> = [
        ['ACCOUNT', 'TOTAL ASSIGN HOURS'],
        [userId, `${totalAssignMinutes/60}`],
    ];
    targetSheet.getRange(`A1:B2`).setValues(cell);

    // TODO: Make the group process ...
    const members: Array<CaAT.IGroupMember> = [{
        name: userId,
        match: new RegExp(null),
    }];
    const holidayWords: CaAT.IHolidayWords = {
        morning: new RegExp(null),
        afternoon: new RegExp(null),
        all: new RegExp(null),
    };
    const groupConfig: CaAT.IGroupConfig = {
        startDate: monday,
        endDate: friday,
        members: members,
        holidayWords: holidayWords,
    };
    const group: CaAT.IGroup = new CaAT.Group(userId, groupConfig);
    const holidays: Array<CaAT.IHoliday> = group.fetchHolidays();
}