import * as CaAT from '@silverbirder/caat'
import {getLatestFriday, getLatestMonday} from './dateUtils';
import {backupSheet, getLastBottomRow, getTemplateSheet, getValueByColumn, ILocation} from './sheet';
import {getSchedules} from './member';
import {getHolidays} from './group';
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const BASE_COL = 'A';
const BASE_ROW = 2;

export interface IMember extends ILocation {
    schedules?: Array<CaAT.ISchedule>,
    holidays?: Array<CaAT.IHoliday>,
}

function main() {
    let _: any;
    let templateSheet: Sheet;
    try {
        _ = backupSheet();
        templateSheet = getTemplateSheet();
    } catch (e) {
        throw e;
    }
    const lastBottomRow: number = getLastBottomRow(templateSheet, `${BASE_COL}${BASE_ROW}`);
    let members: Array<ILocation> = getValueByColumn(templateSheet, `${BASE_COL}${BASE_ROW}:${BASE_COL}${lastBottomRow}`);

    const nextMonday: Date = new Date(getLatestMonday().setHours(0, 0, 0));
    const nextFriday: Date = new Date(getLatestFriday().setHours(23, 59, 59));

    members = getSchedules(members, nextMonday, nextFriday);
    members = getHolidays(members, nextMonday, nextFriday);

    members.forEach((member: IMember) => {
    });
}