import * as XLSX from 'xlsx';
import { BankRow, ConfigMap, ProcessingResult } from '../types';

// Constants
const STANDING_CODES = new Set([469, 515]);
const OVRC_CODES = new Set([120, 175]);
const TRANSFER_CODE = 485;
const TRANSFER_PHRASE = "העב' במקבץ-נט";
const RULE4_CODE = 493;
const RULE4_EPS = 0.50;
const RULE3_AMOUNT_EPS = 0.00;

const RULE5_CODES = new Set([453, 472, 473, 124]);
const RULE6_COMPANY = 'פאיימי בע"מ';
const RULE7_CODE = 143;
const RULE7_PHRASE = "שיקים ממשמרת";
const RULE8_CODE = 191;
const RULE8_PHRASE = "הפק' שיק-שידור";
const RULE9_CODE = 205;
const RULE9_PHRASE = "הפק.שיק במכונה";
const RULE10_CODES = new Set([191, 132, 396]);

// Column Mappings
const COL_MAP = {
  MATCH: ["מס.התאמה", "מס. התאמה", "מס התאמה", "מספר התאמה", "התאמה"],
  BANK_CODE: ["קוד פעולת בנק", "קוד פעולה", "קוד פעולת", "Bank Code"],
  BANK_AMT: ["סכום בדף", "סכום דף", "סכום בבנק", "סכום תנועת בנק", "Bank Amount"],
  BOOKS_AMT: ["סכום בספרים", "סכום בספר", "סכום ספרים", "Books Amount"],
  REF1: ["אסמכתא 1", "אסמכתא1", "אסמכתא", "אסמכתה", "Ref1"],
  REF2: ["אסמכתא 2", "אסמכתא2", "אסמכתא-2", "אסמכתה 2", "Ref2"],
  DATE: ["תאריך מאזן", "תאריך ערך", "תאריך", "Date"],
  DETAILS: ["פרטים", "תיאור", "שם ספק", "Details", "תאור"]
};

const AUX_MAP = {
  DATE: ["תאריך פריקה", "תאריך", "פריקה"],
  AMT: ["אחרי ניכוי", "אחרי", "סכום"],
  PAY: ["מס' תשלום", "מס תשלום", "מספר תשלום"]
};

// Helpers
function pickCol(row: any, candidates: string[]): string | null {
  const keys = Object.keys(row);
  for (const c of candidates) {
    if (keys.includes(c)) return c;
  }
  for (const c of candidates) {
    for (const k of keys) {
      if (k.includes(c)) return k;
    }
  }
  return null;
}

function parseNum(val: any): number {
  if (val === null || val === undefined) return NaN;
  if (typeof val === 'number') return val;
  const str = String(val).replace(/,/g, '').replace(/₪/g, '').replace(/\u200f/g, '').replace(/\u200e/g, '').trim();
  const num = parseFloat(str);
  return isNaN(num) ? NaN : num;
}

function parseDate(val: any): string | null {
  if (!val) return null;
  // Excel date serial handling is done by XLSX usually, but if raw
  if (val instanceof Date) return val.toISOString().split('T')[0];
  // Simple string parsing if needed, assuming XLSX handles most
  return String(val); // Simplification: we rely on XLSX cell dates being standard or strings
}

function onlyDigits(val: any): string {
  return String(val).replace(/\D/g, '').replace(/^0+/, '') || "0";
}

function normalizeDateStr(val: any): string {
    if (!val) return "";
    // If it's a serial number (Excel)
    if (typeof val === 'number') {
        const date = XLSX.SSF.parse_date_code(val);
        // Format to YYYY-MM-DD
        return `${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`;
    }
    // If it's a string, try to parse
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
         return d.toISOString().split('T')[0];
    }
    return String(val);
}


export const processWorkbook = (
  mainFile: ArrayBuffer,
  auxFile: ArrayBuffer | null,
  config: ConfigMap
): ProcessingResult => {
  const wb = XLSX.read(mainFile, { type: 'array', cellDates: true });
  const ws = wb.Sheets["DataSheet"] || wb.Sheets[wb.SheetNames[0]];
  // Use header:1 to get array of arrays, then convert to objects carefully or use standard sheet_to_json
  // sheet_to_json is safer for column name mapping
  let data: BankRow[] = XLSX.utils.sheet_to_json(ws, { defval: "" });

  // Identify Columns
  const sample = data[0] || {};
  const cMatch = pickCol(sample, COL_MAP.MATCH) || "התאמה";
  const cCode = pickCol(sample, COL_MAP.BANK_CODE);
  const cBAmt = pickCol(sample, COL_MAP.BANK_AMT);
  const cAAmt = pickCol(sample, COL_MAP.BOOKS_AMT);
  const cRef1 = pickCol(sample, COL_MAP.REF1);
  const cRef2 = pickCol(sample, COL_MAP.REF2);
  const cDate = pickCol(sample, COL_MAP.DATE);
  const cDet = pickCol(sample, COL_MAP.DETAILS);

  // Initialize Match col if missing
  data.forEach(row => {
    if (row[cMatch] === undefined || row[cMatch] === "") row[cMatch] = 0;
  });

  // --- Rule 1: OV/RC 1:1 ---
  const bankMap = new Map<string, number[]>();
  const bookMap = new Map<string, number[]>();

  data.forEach((row, idx) => {
    if (row[cMatch] !== 0) return;
    const bAmt = parseNum(row[cBAmt]);
    const aAmt = parseNum(row[cAAmt]);
    const date = normalizeDateStr(row[cDate]);

    // Bank Side
    if (cCode && cBAmt && OVRC_CODES.has(parseNum(row[cCode])) && bAmt < 0 && date) {
      const key = `${Math.abs(bAmt).toFixed(2)}_${date}`;
      if (!bankMap.has(key)) bankMap.set(key, []);
      bankMap.get(key)!.push(idx);
    }

    // Books Side
    if (cAAmt && cRef1 && aAmt > 0 && date && String(row[cRef1]).toUpperCase().match(/^(OV|RC)/)) {
      const key = `${Math.abs(aAmt).toFixed(2)}_${date}`;
      if (!bookMap.has(key)) bookMap.set(key, []);
      bookMap.get(key)!.push(idx);
    }
  });

  // Match Rule 1
  for (const [key, bIndices] of bankMap.entries()) {
    const aIndices = bookMap.get(key);
    if (aIndices && bIndices.length === 1 && aIndices.length === 1) {
        // Double check they are still 0 (defensive)
        if(data[bIndices[0]][cMatch] === 0 && data[aIndices[0]][cMatch] === 0) {
            data[bIndices[0]][cMatch] = 1;
            data[aIndices[0]][cMatch] = 1;
        }
    }
  }

  // --- Rule 2: Standing Orders ---
  data.forEach(row => {
    if (row[cMatch] === 0 && cCode && STANDING_CODES.has(parseNum(row[cCode]))) {
      row[cMatch] = 2;
    }
  });

  // --- Rule 3: Transfers (Aux) ---
  const rule3Mismatches: any[] = [];
  if (auxFile) {
      const auxWb = XLSX.read(auxFile, { type: 'array', cellDates: true });
      const auxWs = auxWb.Sheets[auxWb.SheetNames[0]];
      const auxData: any[] = XLSX.utils.sheet_to_json(auxWs);

      const aDateCol = pickCol(auxData[0] || {}, AUX_MAP.DATE);
      const aAmtCol = pickCol(auxData[0] || {}, AUX_MAP.AMT);
      const aPayCol = pickCol(auxData[0] || {}, AUX_MAP.PAY);

      if (aDateCol && aAmtCol) {
          // Group Aux by date
          const auxGroups = new Map<string, { sum: number, pays: Set<string> }>();

          auxData.forEach(r => {
              const dt = normalizeDateStr(r[aDateCol]);
              if (!dt) return;
              const amt = parseNum(r[aAmtCol]);
              if (isNaN(amt)) return;

              if (!auxGroups.has(dt)) auxGroups.set(dt, { sum: 0, pays: new Set() });
              const g = auxGroups.get(dt)!;
              g.sum += amt;
              if (aPayCol && r[aPayCol]) g.pays.add(String(r[aPayCol]).trim());
          });

          // Process groups
          for (const [evtDate, group] of auxGroups.entries()) {
             const evtSum = parseFloat(group.sum.toFixed(2));
             const payset = group.pays;

             // Find Book rows (Ref1 in payset)
             const bookIndices: number[] = [];
             let bookSum = 0;
             if (payset.size > 0 && cRef1 && cAAmt) {
                 data.forEach((row, idx) => {
                     if (row[cMatch] === 0 && payset.has(String(row[cRef1]).trim())) {
                         bookIndices.push(idx);
                         bookSum += parseNum(row[cAAmt]) || 0;
                     }
                 });
             }

             // Find Bank rows (Code 485, phrase match, total sum close to evtSum)
             // This logic in python matches *all* rows that sum up to target?
             // No, the python logic finds rows where: match=0, code=485, amt>0, details contains phrase,
             // AND absolute difference between row amount and evtSum is small.
             // Wait, the python logic: `bank_mask & (bamt.abs().sub(abs(evt_sum)).abs() <= EPS)`
             // This implies 1-to-N matching where a single bank line matches the sum of the aux group.
             const bankIndices: number[] = [];
             if (cCode && cBAmt && cDet) {
                 data.forEach((row, idx) => {
                     if (row[cMatch] === 0 &&
                         parseNum(row[cCode]) === TRANSFER_CODE &&
                         parseNum(row[cBAmt]) > 0 &&
                         String(row[cDet]).includes(TRANSFER_PHRASE)
                        ) {
                            const bVal = parseNum(row[cBAmt]);
                            if (Math.abs(Math.abs(bVal) - Math.abs(evtSum)) <= RULE3_AMOUNT_EPS) {
                                bankIndices.push(idx);
                            }
                        }
                 });
             }

             const bookSumFixed = parseFloat(bookSum.toFixed(2));
             const absDiff = Math.abs(Math.abs(bookSumFixed) - Math.abs(evtSum));

             if (bankIndices.length > 0 && bookIndices.length > 0) {
                 if (absDiff <= RULE3_AMOUNT_EPS) {
                     // Match!
                     bankIndices.forEach(i => data[i][cMatch] = 3);
                     bookIndices.forEach(i => data[i][cMatch] = 3);
                 } else {
                     rule3Mismatches.push({
                         "Event": evtDate,
                         "Aux Sum": evtSum,
                         "Books Sum": bookSumFixed,
                         "Gap": absDiff.toFixed(2),
                         "Bank Count": bankIndices.length,
                         "Books Count": bookIndices.length
                     });
                 }
             } else {
                 rule3Mismatches.push({
                     "Event": evtDate,
                     "Aux Sum": evtSum,
                     "Books Sum": bookIndices.length ? bookSumFixed : "N/A",
                     "Gap": "N/A",
                     "Bank Count": bankIndices.length,
                     "Books Count": bookIndices.length
                 });
             }
          }
      }
  }

  // --- Rule 4: Checks ---
  // Bank: Code 493, Ref1 exists. Books: Ref1 starts "CH", Ref2 exists.
  // Match Ref1(digits) == Ref2(digits) and abs(diff) <= 0.50
  if (cCode && cRef1 && cBAmt && cAAmt && cRef2) {
      const bank4 = data.map((r, i) => ({r, i})).filter(item =>
          item.r[cMatch] === 0 &&
          parseNum(item.r[cCode]) === RULE4_CODE &&
          String(item.r[cRef1]).trim() !== "" &&
          !isNaN(parseNum(item.r[cBAmt]))
      );

      const books4 = data.map((r, i) => ({r, i})).filter(item =>
          item.r[cMatch] === 0 &&
          String(item.r[cRef1]).toUpperCase().startsWith("CH") &&
          String(item.r[cRef2]).trim() !== "" &&
          !isNaN(parseNum(item.r[cAAmt]))
      );

      const usedBooks = new Set<number>();

      for (const bItem of bank4) {
          const bRefClean = onlyDigits(bItem.r[cRef1]);
          const bAmt = Math.abs(parseNum(bItem.r[cBAmt]));

          for (const aItem of books4) {
              if (usedBooks.has(aItem.i) || aItem.r[cMatch] !== 0) continue;

              const aRefClean = onlyDigits(aItem.r[cRef2]);
              if (bRefClean !== aRefClean) continue;

              const aAmt = Math.abs(parseNum(aItem.r[cAAmt]));
              if (Math.abs(aAmt - bAmt) <= RULE4_EPS) {
                  data[bItem.i][cMatch] = 4;
                  data[aItem.i][cMatch] = 4;
                  usedBooks.add(aItem.i);
                  break;
              }
          }
      }
  }

  // --- Rules 5-10 ---
  data.forEach(row => {
      if (row[cMatch] !== 0) return;
      const code = parseNum(row[cCode]);
      const bAmt = parseNum(row[cBAmt]);
      const det = String(row[cDet] || "");

      // 5
      if (RULE5_CODES.has(code) && bAmt > 0 && bAmt <= 1000) {
          row[cMatch] = 5;
          return;
      }
      // 6
      if (code === 175 && bAmt < 0 && det === RULE6_COMPANY) {
          row[cMatch] = 6;
          return;
      }
      // 7
      if (code === RULE7_CODE && bAmt < 0 && det === RULE7_PHRASE) {
          row[cMatch] = 7;
          return;
      }
      // 8
      if (code === RULE8_CODE && bAmt < 0 && det === RULE8_PHRASE) {
          row[cMatch] = 8;
          return;
      }
      // 9
      if (code === RULE9_CODE && bAmt < 0 && det === RULE9_PHRASE) {
          row[cMatch] = 9;
          return;
      }
      // 10
      if (RULE10_CODES.has(code) && !isNaN(bAmt) && bAmt !== 0) {
          row[cMatch] = 10;
          return;
      }
  });

  // --- Rule 11: BT (Placeholder replacement) ---
  // Match 0, Code 485 vs Ref1 starts "BT". 1:1 by absolute amount.
  if (cCode && cBAmt && cAAmt && cRef1) {
      const bank11Map = new Map<string, number[]>();
      const books11Map = new Map<string, number[]>();

      // Collect candidates
      data.forEach((row, idx) => {
          if (row[cMatch] !== 0) return;
          const code = parseNum(row[cCode]);
          const bAmt = parseNum(row[cBAmt]);
          const aAmt = parseNum(row[cAAmt]);
          const ref1 = String(row[cRef1]).trim();

          // Bank candidate
          if (code === 485 && bAmt !== 0 && !isNaN(bAmt)) {
              const key = Math.abs(bAmt).toFixed(2);
              if (!bank11Map.has(key)) bank11Map.set(key, []);
              bank11Map.get(key)!.push(idx);
          }

          // Books candidate
          if (aAmt !== 0 && !isNaN(aAmt) && ref1.toUpperCase().startsWith("BT")) {
              const key = Math.abs(aAmt).toFixed(2);
              if (!books11Map.has(key)) books11Map.set(key, []);
              books11Map.get(key)!.push(idx);
          }
      });

      // Match
      for (const [key, bIdxList] of bank11Map.entries()) {
          const aIdxList = books11Map.get(key);
          if (!aIdxList) continue;

          const count = Math.min(bIdxList.length, aIdxList.length);
          for (let k = 0; k < count; k++) {
              data[bIdxList[k]][cMatch] = 11;
              data[aIdxList[k]][cMatch] = 11;
          }
      }
  }

  // --- VLOOKUP Sheet Generation ---
  const vlookupRows: any[] = [];
  let totalCreditWithSupplier = 0;

  data.forEach(row => {
      if (row[cMatch] === 2) {
          const det = String(row[cDet] || "");
          const amt = parseNum(row[cBAmt]);
          let supplier = "";

          // Name Map
          for (const [k, v] of Object.entries(config.nameMap)) {
              if (k && det.includes(k)) {
                  supplier = v;
                  break;
              }
          }

          // Amount Map fallback
          if (!supplier && !isNaN(amt)) {
              const k = Math.abs(amt).toFixed(2);
              if (config.amountMap[k]) supplier = config.amountMap[k];
          }

          const hova = !isNaN(amt) ? Math.abs(amt) : 0;
          if (supplier) totalCreditWithSupplier += hova;

          vlookupRows.push({
              "פרטים": det,
              "סכום": amt,
              "מס' ספק": supplier,
              "סכום חובה": hova,
              "סכום זכות": 0
          });
      }
  });

  if (totalCreditWithSupplier > 0) {
      vlookupRows.push({
          "פרטים": 'סה"כ זכות – עם מס\' ספק',
          "סכום": 0,
          "מס' ספק": 20001,
          "סכום חובה": 0,
          "סכום זכות": parseFloat(totalCreditWithSupplier.toFixed(2))
      });
  }

  // Calc Stats
  const counts: Record<number, number> = {};
  data.forEach(r => {
      const m = r[cMatch] || 0;
      counts[m] = (counts[m] || 0) + 1;
  });
  const stats = Object.keys(counts).map(k => ({ rule: k, count: counts[parseInt(k)] })).sort((a,b) => Number(a.rule) - Number(b.rule));

  return {
      processedData: data,
      stats,
      vlookupData: vlookupRows,
      rule3Mismatches
  };
};
