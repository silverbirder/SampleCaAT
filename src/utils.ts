export function getLatestMonday() :Date{
    const now: Date = new Date();
    return new Date(now.setDate(now.getDate() - (now.getDay() + 6) % 7 + 7));
}

export function getLatestFriday(): Date{
    const now: Date = new Date();
    return new Date(now.setDate(now.getDate() - (now.getDay() + 2) % 7 + 7));
}