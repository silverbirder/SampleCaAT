import * as CaAT from '@silverbirder/caat';
import {copyDate} from './dateUtils';
import {IMember} from "./index";
import {IMemberConfig} from "@silverbirder/caat";

export function getSchedules(members: Array<IMember>, startDate: Date, endDate: Date): Array<IMember> {
    const cutTimeRange: Array<CaAT.IRange> = [];
    for (const d = copyDate(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        cutTimeRange.push({
            from: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0),
            to: new Date(d.getFullYear(), d.getMonth(), d.getDate(), 13, 0, 0),
        })
    }
    const memberConfig: CaAT.IMemberConfig = {
        startDate: startDate,
        endDate: endDate,
        everyMinutes: 15,
        cutTimeRange: cutTimeRange,
        ignore: new RegExp(null),
    };

    // Fetch the schedules!
    return members.map((member: IMember) => {
        const cMemberConfig: IMemberConfig = Object.create(memberConfig);
        if (member.holidays) {
            const range: Array<CaAT.IRange> = member.holidays.map((holiday: CaAT.IHoliday) => {
                const t: Date = holiday.toDate;
                if (holiday.all) {
                    return {
                        from: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 0, 0, 0),
                        to: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 23, 59, 59),
                    }
                } else if (holiday.morning) {
                    return {
                        from: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 0, 0, 0),
                        to: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 12, 0, 0),
                    }
                } else {
                    return {
                        from: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 12, 0, 0),
                        to: new Date(t.getFullYear(), t.getMonth(), t.getDate(), 23, 59, 59),
                    }
                }
            });
            cMemberConfig.cutTimeRange = cMemberConfig.cutTimeRange.concat(range);
        }
        const caatMember: CaAT.IMember = new CaAT.Member(`${member.value}@gmail.com`, cMemberConfig);
        member.schedules = caatMember.fetchSchedules();
        return member;
    })
}