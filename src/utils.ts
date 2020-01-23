export function getLatestMonday() :Date{
    const now: Date = new Date();
    return new Date(now.setDate(now.getDate() - (now.getDay() + 6) % 7 + 7));
}

export function skipDay(d: Date, skip: number): Date{
    const copy: Date = copyDate(d);
    copy.setDate(copy.getDate()+skip);
    return copy;
}

export function copyDate(d: Date) {
    return new Date(d.getTime());
}