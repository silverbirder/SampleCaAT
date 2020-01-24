export function getLatestMonday(): Date {
    const previousMonday: Date = getPreviousMonday();
    return skipDay(previousMonday, 7);
}

export function getPreviousMonday(): Date {
    const now: Date = new Date();
    return new Date(now.setDate(now.getDate() - (now.getDay() + 6) % 7));
}

export function getLatestFriday(): Date {
    const latestMonday: Date = getLatestMonday();
    return skipDay(latestMonday, 4);
}

export function skipDay(d: Date, skip: number): Date {
    const copy: Date = copyDate(d);
    copy.setDate(copy.getDate() + skip);
    return copy;
}

export function copyDate(d: Date) {
    return new Date(d.getTime());
}