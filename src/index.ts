import CaAT from '@silverbirder/caat'

function getMembers(){
    const id: string = 'testUser@gmail.com';
    const config: CaAT.IMemberConfig = {
        startDate: new Date(),
        endDate: new Date(),
        everyMinutes: 15,
        cutTimeRange: [],
        ignore: new RegExp(''),
    }
    const member: CaAT.IMember = new CaAT.Member(id, config);
    const schedules: Array<CaAT.ISchedule> = member.fetchSchedules();
    Logger.log(schedules);
}