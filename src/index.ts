import * as CaAT from '@silverbirder/caat'
import {copyDate, getLatestMonday, skipDay} from "./utils";

function getMembers() {
    const id: string = PropertiesService.getScriptProperties().getProperty('ID');

    // setup date
    const cutTimeRange: Array<CaAT.IRange> = [];
    const monday: Date = new Date(getLatestMonday().setHours(0, 0, 0));
    const friday: Date = new Date(skipDay(monday, 4).setHours(23, 59, 59));
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
        ignore: new RegExp(''),
    };
    const member: CaAT.IMember = new CaAT.Member(id, config);
    const schedules: Array<CaAT.ISchedule> = member.fetchSchedules();
}