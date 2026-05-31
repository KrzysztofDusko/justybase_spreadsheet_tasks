export const F = {
    THOUSANDS_SEP: '#,##0',
    CURRENCY_PLN: '#,##0.00 "z\u0142"',
    CURRENCY_EUR: '#,##0.00 \u20AC',
    PERCENTAGE: '0%',
    SCIENTIFIC: '0.00E+00',
    TWO_DECIMALS: '#,##0.00',
    TEXT: '@',
    LEADING_ZEROS: '000000000',

    DATE_SHORT: 'dd.mm.yyyy',
    DATE_LONG: 'd mmmm yyyy',
    DATE_DAY_MONTH_YEAR: 'dd-mm-yyyy',
    DATE_ISO: 'yyyy-mm-dd',
    DATE_MONTH_YEAR: 'mmmm yyyy',
    DATE_WEEKDAY: 'dddd, d mmmm yyyy',
    DATE_DAY_MONTH: 'd mmmm',
    DATE_YEAR_ONLY: 'yyyy',

    DATETIME_SHORT: 'dd.mm.yyyy hh:mm',
    DATETIME_LONG: 'd mmmm yyyy hh:mm:ss',
    TIME_HH_MM: 'hh:mm',
    TIME_HH_MM_SS: 'hh:mm:ss',
    TIME_12H: 'h:mm AM/PM',
    DATETIME_24H: 'dd.mm.yyyy hh:mm:ss',
    DATETIME_ISO: 'yyyy-mm-dd"T"hh:mm:ss',
    TIME_MS: 'hh:mm:ss.000',
};

export function isFormattedCell(val) {
    return val !== null && val !== undefined && typeof val === 'object' && 'format' in val && typeof val.format === 'string';
}

export function unwrapCell(val) {
    return isFormattedCell(val) ? val.value : val;
}

export function getFormat(val) {
    return isFormattedCell(val) ? val.format : null;
}
