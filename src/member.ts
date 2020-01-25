import * as CaAT from '@silverbirder/caat';
import {copyDate} from './dateUtils';
import {IMember} from "./index";

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
        const caatMember: CaAT.IMember = new CaAT.Member(`${member.value}@gmail.com`, memberConfig);
        member.schedules = caatMember.fetchSchedules();
        return member;
    })
}