---
"@betteroffice/xlsx": patch
---

Date functions (`DATE`, `DAY`, `MONTH`, `YEAR`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`, `EDATE`, `EOMONTH`, `DAYS`, `YEARFRAC`, `NETWORKDAYS`, `WORKDAY`) return `#NUM!` for dates past 9999-12-31 or huge day counts, instead of a wrong date or an overflow.
