# Адема в Японии

Travel itinerary site for Adema's Japan tours — March 2025 + July 2026.

**Live**: https://elbicsabitov.github.io/adema-japan-tour/

Single-file design: open `index.html` and it works offline (data comes from `tour_data.min.js`).

## Source data

- `Тур для Адемы v3.xlsx` → March 2025 tour (Fukuoka → Hiroshima → Osaka → Nara → Hakone → Tokyo)
- `Адема Июль 2026 (draft ver_5) (3).xlsx` → July 2026 tour (Kansai → Uji → Kyoto → Takashima → Hikone → Nagoya → Shizuoka → Fujikawaguchiko → Tokyo)

**Coverage**: 99% of source rows (151/152 march events, 67/68 july events, 100% of hotels, 100% of flights, 100% of attractions). Budget components sum to $3297.88 vs declared $3297.87.

**Links**: ALL URLs are real, extracted directly from Excel via openpyxl `cell.hyperlink.target` (not invented). Booking, Klook, official temple sites, Google Maps — all from source file.

| March | Real URLs |
|---|---|
| Map links | 129 |
| Site links | 33 |
| Hotel bookings | 10 (incl. 2 alternatives for Kurashiki) |

| July | Real URLs |
|---|---|
| Map links | 63 |
| Site links | 21 |
| Hotel bookings | 22 |
| Attraction sites | 39 |

Re-extract from Excel: `python build_data.py`

## Design references

Editorial wabi-sabi direction informed by:

- Edward Tufte — small multiples + data-ink ratio (budget breakdown bar)
- Massimo Vignelli — typographic hierarchy by scale, not weight
- Kenya Hara (MUJI) — emptiness as substance
- Dieter Rams — less but better
- Don Norman — affordances (map links visually obvious)
- Linda Dong (Apple HIG) — sticky day rail, restraint
- Erik Spiekermann — tabular figures for prices/times

Typography: Cormorant Garamond (display, Cyrillic) + Onest (sans, Cyrillic) + Noto Serif JP (Japanese accents) + JetBrains Mono (numbers).

Colour: washi paper + sumi ink + torii vermillion + ki-iro gold + asagi indigo.
