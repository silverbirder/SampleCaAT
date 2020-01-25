import * as CaAT from '@silverbirder/caat'
import {IHoliday} from '@silverbirder/caat'
import {IMember} from "./index";

export function getHolidays(members: Array<IMember>, startDate: Date, endDate: Date): Array<IMember> {
    const groupMembers: Array<CaAT.IGroupMember> = [];
    members.forEach((member: IMember) => {
        const name: string = member.value;
        const displayName: string = PropertiesService.getScriptProperties().getProperty(`MEMBER_${name}`);
        groupMembers.push({
            name: name,
            match: new RegExp(displayName !== '' ? displayName : null),
        })
    });
    const holidayWords: CaAT.IHolidayWords = {
        all: new RegExp('休'),
        morning: new RegExp('AM|午前'),
        afternoon: new RegExp('PM|午後'),
    };

    const groupConfig: CaAT.IGroupConfig = {
        startDate: startDate,
        endDate: endDate,
        members: groupMembers,
        holidayWords: holidayWords,
    };

    const groupId = PropertiesService.getScriptProperties().getProperty('GROUP_ID');
    const caatGroup: CaAT.IGroup = new CaAT.Group(groupId, groupConfig);
    const holidays: Array<IHoliday> = caatGroup.fetchHolidays();
    for (let m = 0; m < members.length; m++) {
        members[m].holidays = [];
        for (let h = 0; h < holidays.length; h++) {
            if (holidays[h].inMember.indexOf(members[m].value) !== -1) {
                members[m].holidays.push(holidays[h]);
            }
        }
    }
    return members;
}