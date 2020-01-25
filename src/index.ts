import * as CaAT from '@silverbirder/caat'
import {copyDate, getLastBottomRow, getLatestFriday, getLatestMonday} from "./utils";
import {backupSheet, getValueByColumn, getTemplateSheet, ILocation} from "./sheet";
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const BASE_COL = 'A';
const BASE_ROW = 2;
const ASSIGN_COL = 'G';

function main() {
    let templateSheet: Sheet;
    try {
        backupSheet();
        templateSheet = getTemplateSheet();
    } catch (e) {
        throw e;
    }

    // Prepare CaAT member config
    const nextMonday: Date = new Date(getLatestMonday().setHours(0, 0, 0));
    const nextFriday: Date = new Date(getLatestFriday().setHours(23, 59, 59));
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
    const members: Array<ILocation> = getValueByColumn(templateSheet, `${BASE_COL}${BASE_ROW}:${BASE_COL}${lastBottomRow}`);

    // Fetch the schedules!
    members.forEach((member: ILocation) => {
        const caatMember: CaAT.IMember = new CaAT.Member(`${member.value}@gmail.com`, config);
        const schedules: Array<CaAT.ISchedule> = caatMember.fetchSchedules();
        const totalAssignMinutes: number = schedules.reduce((totalAssignTime: number, schedule: CaAT.ISchedule): number => {
            return totalAssignTime + schedule.assignMinute;
        }, 0);
        templateSheet.getRange(`${ASSIGN_COL}${member.row}`).setFormula(`=CEILING(${totalAssignMinutes / 60}, 0.25)`);
    })
}