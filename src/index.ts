import * as CaAT from '@silverbirder/caat'
import {copyDate, getLatestFriday, getLatestMonday} from './dateUtils';
import {
    backupSheet,
    COL,
    getColList,
    getLastBottomRow,
    getRowList,
    getTemplateSheet,
    getValueByColumn,
    ILocation,
    initSheet,
    ROW
} from './sheet';
import {getSchedules} from './member';
import {getHolidays} from './group';
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const FULL_WORK_TIME_HOUR = 5;
const MORNING_WORK_TIME_HOUR = 2;
const AFTERNOON_WORK_TIME_HOUR = 3;

export interface IMember extends ILocation {
    schedules?: Array<CaAT.ISchedule>,
    holidays?: Array<CaAT.IHoliday>,
}

function main() {
    const _: Sheet = backupSheet();
    const templateSheet: Sheet = getTemplateSheet();

    const col: COL = getColList();
    const row: ROW = getRowList();
    const lastBottomRow: number = getLastBottomRow(templateSheet, `${col.base}${row.base}`);
    const nextMonday: Date = getLatestMonday();
    const initMonday: Date = new Date(nextMonday.getFullYear(), nextMonday.getMonth(), nextMonday.getDate(), 0,0,0);
    const nextFriday: Date = getLatestFriday();
    const initFriday: Date = new Date(nextFriday.getFullYear(), nextFriday.getMonth(), nextFriday.getDate(), 23, 59, 59);
    let members: Array<ILocation> = getValueByColumn(templateSheet, `${col.base}${row.base}:${col.base}${lastBottomRow}`);

    initSheet(templateSheet, initMonday, FULL_WORK_TIME_HOUR, lastBottomRow);

    members = getHolidays(members, initMonday, initFriday);
    members = getSchedules(members, initMonday, initFriday);

    members.forEach((member: IMember) => {
        const totalAssignMinutes: number = member.schedules.reduce((totalAssignMinutes: number, schedule: CaAT.ISchedule) => {
            return totalAssignMinutes + schedule.assignMinute;
        }, 0);
        member.holidays.forEach((holiday: CaAT.IHoliday) => {
            let reduceHour;
            if (holiday.all) {
                reduceHour = FULL_WORK_TIME_HOUR;
            } else if (holiday.morning) {
                reduceHour = MORNING_WORK_TIME_HOUR;
            } else {
                reduceHour = AFTERNOON_WORK_TIME_HOUR;
            }
            let targetCol = '';
            switch (holiday.toDate.getDay()) {
                case 1:
                    targetCol = col.mon;
                    break;
                case 2:
                    targetCol = col.tue;
                    break;
                case 3:
                    targetCol = col.wed;
                    break;
                case 4:
                    targetCol = col.thu;
                    break;
                case 5:
                    targetCol = col.fri;
                    break;
            }
            if (targetCol != '') {
                const remainHour: number = parseInt(templateSheet.getRange(`${targetCol}${member.row}`).getValue());
                templateSheet.getRange(`${targetCol}${member.row}`).setFormula(`=CEILING(${remainHour - reduceHour}, 0.25)`);
            }
        });
        templateSheet.getRange(`${col.assign}${member.row}`).setFormula(`=CEILING(${totalAssignMinutes / 60}, 0.25)`);
    });
}