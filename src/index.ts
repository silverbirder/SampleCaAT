import * as CaAT from '@silverbirder/caat'
import {getLatestFriday, getLatestMonday} from './dateUtils';
import {backupSheet, getLastBottomRow, getTemplateSheet, getValueByColumn, ILocation} from './sheet';
import {getSchedules} from './member';
import {getHolidays} from './group';
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const BASE_COL = 'A';
const BASE_ROW = 2;

const MON_COL = 'B';
const TUE_COL = 'C';
const WED_COL = 'D';
const THU_COL = 'E';
const FRI_COL = 'F';

const ASSIGN_COL = 'G';

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

    members = getHolidays(members, nextMonday, nextFriday);
    members = getSchedules(members, nextMonday, nextFriday);

    members.forEach((member: IMember) => {
        const totalAssignMinutes: number = member.schedules.reduce((totalAssignMinutes: number, schedule: CaAT.ISchedule) => {
            return totalAssignMinutes + schedule.originalAssignMinute;
        }, 0);
        member.holidays.forEach((holiday: CaAT.IHoliday) => {
            let reduceHour = 0;
            if (holiday.all) {
                reduceHour = 5;
            } else {
                reduceHour = 2.5;
            }
            let targetCol = '';
            switch (holiday.toDate.getDay()) {
                case 1:
                    targetCol = MON_COL;
                    break;
                case 2:
                    targetCol = TUE_COL;
                    break;
                case 3:
                    targetCol = WED_COL;
                    break;
                case 4:
                    targetCol = THU_COL;
                    break;
                case 5:
                    targetCol = FRI_COL;
                    break;
            }
            const remainHour: number = parseInt(templateSheet.getRange(`${targetCol}${member.row}`).getValue());
            templateSheet.getRange(`${targetCol}${member.row}`).setFormula(`=CEILING(${remainHour - reduceHour}, 0.25)`);
        });
        templateSheet.getRange(`${ASSIGN_COL}${member.row}`).setFormula(`=CEILING(${totalAssignMinutes / 60}, 0.25)`);
    });
}