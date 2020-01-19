function sample() {
    // @ts-ignore
    const member = new CaAT.MemberImpl('testUser@gmail.com');
    member.ignore = new RegExp('Concentration', 'i');
    member.everyMinutes = 15;
    member.cutTimeRange = [{from: new Date('2020-01-18T12:00'), to: new Date('2020-01-18T13:00')}];
    member.startDate = new Date('2020-01-20T09:00:00');
    member.endDate = new Date('2020-01-24T18:00:00');
    const schedules  = member.fetchSchedules();
}
