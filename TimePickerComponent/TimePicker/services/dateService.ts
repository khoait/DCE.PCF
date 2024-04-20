export function getUtcDate(localDate: Date | null) {
  if (!localDate) return null;

  return new Date(
    localDate.getUTCFullYear(),
    localDate.getUTCMonth(),
    localDate.getUTCDate(),
    localDate.getUTCHours(),
    localDate.getUTCMinutes()
  );
}

export function convertDate(value: Date | null, timezoneOffsetMinutes: number) {
  if (!value) return null;

  const localDate = addMinutes(value, timezoneOffsetMinutes);
  return getUtcDate(localDate);
}

export function addMinutes(date: Date, minutes: number): Date {
  return new Date(date.getTime() + minutes * 60000);
}

export function isValidDate(date: Date) {
  return date instanceof Date && !isNaN(date.getTime());
}
