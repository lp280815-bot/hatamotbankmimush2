import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileUp, Settings, Download, CheckCircle, AlertTriangle, Play, FileSpreadsheet } from 'lucide-react';
import { processWorkbook } from './excelLogic';
import { ConfigMap, ProcessingResult } from './types';

// Default Config
const DEFAULT_CONFIG: ConfigMap = {
  nameMap: {
    "בזק בינלאומי ב": "30006",
    "פרי ירוחם חב'": "34714",
    "סלקום ישראל בע": "30055",
    "בזק-הוראות קבע": "34746",
    "דרך ארץ הייווי": "34602",
    "גלובס פבלישר ע": "30067",
    "פלאפון תקשורת": "30030",
    "מרכז הכוכביות": "30002",
    "ע.אשדוד-מסים": "30056",
    "א.ש.א(בס\"ד)אחז": "30050",
    "או.פי.ג'י(מ.כ)": "30047",
    "רשות האכיפה וה": "67-1",
    "קול ביז מילניו": "30053",
    "פריוריטי סופטו": "30097",
    "אינטרנט רימון": "34636",
    "עו\"דכנית בע\"מ": "30018",
    "עיריית רמת גן": "30065",
    "פז חברת נפט בע": "34811",
    "ישראכרט": "28002",
    "חברת החשמל ליש": "30015",
    "הפניקס ביטוח": "34686",
    "מימון ישיר מקב": "34002",
    "שלמה טפר": "30247",
    "נמרוד תבור עורך-דין": "30038",
    "עיריית בית שמש": "34805",
    "פז קמעונאות וא": "34811",
    "הו\"ק הלו' רבית": "8004",
  },
  amountMap: {}
};

function App() {
  const [config, setConfig] = useState<ConfigMap>(DEFAULT_CONFIG);
  const [mainFile, setMainFile] = useState<File | null>(null);
  const [auxFile, setAuxFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<ProcessingResult | null>(null);
  const [showConfig, setShowConfig] = useState(false);

  // Config Modal State
  const [newNameKey, setNewNameKey] = useState("");
  const [newNameVal, setNewNameVal] = useState("");
  const [newAmtKey, setNewAmtKey] = useState("");
  const [newAmtVal, setNewAmtVal] = useState("");

  useEffect(() => {
    const saved = localStorage.getItem('bank_recon_config');
    if (saved) {
      try {
        setConfig(JSON.parse(saved));
      } catch(e) { console.error(e); }
    }
  }, []);

  const saveConfig = (newConfig: ConfigMap) => {
    setConfig(newConfig);
    localStorage.setItem('bank_recon_config', JSON.stringify(newConfig));
  };

  const handleProcess = async () => {
    if (!mainFile) return;
    setIsProcessing(true);
    setResult(null);

    // Simulate async for UI responsiveness
    setTimeout(async () => {
        try {
            const mainBuffer = await mainFile.arrayBuffer();
            const auxBuffer = auxFile ? await auxFile.arrayBuffer() : null;

            const res = processWorkbook(mainBuffer, auxBuffer, config);
            setResult(res);
        } catch (error) {
            console.error(error);
            alert("שגיאה בעיבוד הקובץ. אנא ודאי שהקובץ תקין.");
        } finally {
            setIsProcessing(false);
        }
    }, 100);
  };

  const downloadResults = () => {
    if (!result) return;
    const wb = XLSX.utils.book_new();

    // DataSheet
    const wsData = XLSX.utils.json_to_sheet(result.processedData);
    XLSX.utils.book_append_sheet(wb, wsData, "DataSheet");

    // Summary
    const wsStats = XLSX.utils.json_to_sheet(result.stats.map(s => ({ "מס": s.rule, "כמות": s.count })));
    XLSX.utils.book_append_sheet(wb, wsStats, "סיכום");

    // VLOOKUP Sheet (Standing Orders)
    const wsVlookup = XLSX.utils.json_to_sheet(result.vlookupData);
    XLSX.utils.book_append_sheet(wb, wsVlookup, "הוראת קבע ספקים");

    // Rule 3 Gaps
    if (result.rule3Mismatches.length > 0) {
        const wsGaps = XLSX.utils.json_to_sheet(result.rule3Mismatches);
        XLSX.utils.book_append_sheet(wb, wsGaps, "פערי סכומים – כלל 3");
    }

    // Set RTL for all sheets (Workbook property)
    wb.Workbook = { Views: [{ RTL: true }] };

    XLSX.writeFile(wb, "התאמות_1_עד_12.xlsx");
  };

  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["פרטים", "סכום", "מס' ספק"],
      ["שם ספק לדוגמה", "", "12345"],
      ["", 150.50, "98765"]
    ]);
    ws['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, ws, "ספקים");
    XLSX.writeFile(wb, "תבנית_יבוא_ספקים.xlsx");
  };

  const importConfigExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if(!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type: 'array'});
    const data = XLSX.utils.sheet_to_json<any>(wb.Sheets[wb.SheetNames[0]]);
    
    const newNameMap = {...config.nameMap};
    const newAmountMap = {...config.amountMap};
    let countName = 0;
    let countAmount = 0;

    data.forEach(row => {
        const sup = row["מס' ספק"] || row["מס ספק"] || row["Supplier"];
        if (!sup) return;

        // Name Mapping
        const det = row["פרטים"] || row["תיאור"] || row["Details"];
        if(det) {
            newNameMap[String(det).trim()] = String(sup).trim();
            countName++;
        }

        // Amount Mapping
        const amt = row["סכום"] || row["Amount"];
        if (amt !== undefined && amt !== null && amt !== "") {
            const num = parseFloat(String(amt).replace(/,/g, ''));
            if (!isNaN(num)) {
                 const k = Math.abs(num).toFixed(2);
                 newAmountMap[k] = String(sup).trim();
                 countAmount++;
            }
        }
    });
    
    saveConfig({...config, nameMap: newNameMap, amountMap: newAmountMap});
    alert(`נטענו ${countName} ספקים לפי שם ו-${countAmount} לפי סכום.`);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans" dir="rtl">
      {/* Header */}
      <header className="bg-white shadow-sm sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
            <div className="flex items-center gap-3">
                <div className="bg-emerald-600 text-white p-2 rounded-lg shadow-sm">
                    <CheckCircle size={24} />
                </div>
                <h1 className="text-xl font-bold text-slate-800">התאמות בנק – 1 עד 12- מימוש</h1>
            </div>
            <button
                onClick={() => setShowConfig(true)}
                className="flex items-center gap-2 text-slate-500 hover:text-emerald-600 transition-colors"
            >
                <Settings size={20} />
                <span>הגדרות ספקים</span>
            </button>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-8">

        {/* Upload Section */}
        <section className="grid md:grid-cols-2 gap-6">
            <div className={`p-8 rounded-2xl border-2 border-dashed transition-all duration-200 flex flex-col items-center justify-center text-center h-64 bg-white
                ${mainFile ? 'border-emerald-500 bg-emerald-50' : 'border-slate-200 hover:border-emerald-400'}`}>
                <FileSpreadsheet size={48} className={mainFile ? 'text-emerald-600' : 'text-slate-300'} />
                <div className="mt-4">
                    <h3 className="font-semibold text-lg">קובץ מקור (DataSheet)</h3>
                    <p className="text-sm text-slate-500 mt-1">גררי לכאן את הקובץ הראשי</p>
                </div>
                <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => setMainFile(e.target.files?.[0] || null)}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                    title=""
                />
                {mainFile && <span className="mt-4 inline-block bg-white px-3 py-1 rounded-full text-sm font-medium text-emerald-700 shadow-sm">{mainFile.name}</span>}
            </div>

            <div className={`p-8 rounded-2xl border-2 border-dashed transition-all duration-200 flex flex-col items-center justify-center text-center h-64 bg-white
                ${auxFile ? 'border-blue-500 bg-blue-50' : 'border-slate-200 hover:border-blue-400'}`}>
                <FileUp size={48} className={auxFile ? 'text-blue-600' : 'text-slate-300'} />
                <div className="mt-4">
                    <h3 className="font-semibold text-lg">קובץ עזר להעברות</h3>
                    <p className="text-sm text-slate-500 mt-1">אופציונלי (עבור כלל 3)</p>
                </div>
                <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => setAuxFile(e.target.files?.[0] || null)}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full" 
                    style={{display: 'block', width: '100%', height: '100%'}} 
                    title=""
                />
                 {auxFile && <span className="mt-4 inline-block bg-white px-3 py-1 rounded-full text-sm font-medium text-blue-700 shadow-sm">{auxFile.name}</span>}
            </div>
        </section>

        {/* Action Bar */}
        <div className="flex justify-center">
            <button
                onClick={handleProcess}
                disabled={!mainFile || isProcessing}
                className={`flex items-center gap-3 px-8 py-4 rounded-xl shadow-lg text-lg font-bold transition-all transform hover:-translate-y-1
                ${!mainFile || isProcessing ? 'bg-slate-300 text-slate-500 cursor-not-allowed' : 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white hover:shadow-emerald-200'}`}
            >
                {isProcessing ? (
                    <>
                        <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
                        מעבד נתונים...
                    </>
                ) : (
                    <>
                        <Play fill="currentColor" />
                        הרץ התאמות (1-12)
                    </>
                )}
            </button>
        </div>

        {/* Results Section */}
        {result && (
            <div className="animate-fade-in space-y-6">
                <div className="flex items-center justify-between">
                    <h2 className="text-2xl font-bold text-slate-800">תוצאות העיבוד</h2>
                    <button
                        onClick={downloadResults}
                        className="flex items-center gap-2 bg-blue-600 text-white px-5 py-2.5 rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-md"
                    >
                        <Download size={20} />
                        הורד קובץ מעודכן
                    </button>
                </div>

                {/* Stats Grid */}
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
                    {result.stats.map((stat) => (
                        <div key={stat.rule} className="bg-white p-4 rounded-xl shadow-sm border border-slate-100 flex flex-col items-center">
                            <span className="text-slate-400 text-xs font-medium uppercase tracking-wider">כלל {stat.rule}</span>
                            <span className="text-3xl font-bold text-slate-700 mt-1">{stat.count}</span>
                        </div>
                    ))}
                </div>

                {/* Rule 3 Mismatches Alert */}
                {result.rule3Mismatches.length > 0 && (
                    <div className="bg-amber-50 border border-amber-200 rounded-xl p-6">
                        <div className="flex items-start gap-4">
                            <AlertTriangle className="text-amber-500 shrink-0 mt-1" />
                            <div>
                                <h3 className="text-lg font-bold text-amber-800">נמצאו פערים בכלל 3 (העברות)</h3>
                                <p className="text-amber-700 mt-1">ישנם {result.rule3Mismatches.length} אירועים בהם סכום הספרים לא תאם את סכום קובץ העזר.</p>
                                <div className="mt-4 overflow-x-auto">
                                    <table className="w-full text-sm text-right">
                                        <thead>
                                            <tr className="text-amber-900 border-b border-amber-200">
                                                <th className="pb-2">אירוע</th>
                                                <th className="pb-2">סכום עזר</th>
                                                <th className="pb-2">סכום ספרים</th>
                                                <th className="pb-2">פער</th>
                                            </tr>
                                        </thead>
                                        <tbody className="text-amber-800">
                                            {result.rule3Mismatches.slice(0, 5).map((m, i) => (
                                                <tr key={i} className="border-b border-amber-100 last:border-0">
                                                    <td className="py-2 ltr text-right">{m["Event"]}</td>
                                                    <td className="py-2 font-mono">{m["Aux Sum"]}</td>
                                                    <td className="py-2 font-mono">{m["Books Sum"]}</td>
                                                    <td className="py-2 font-mono font-bold">{m["Gap"]}</td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                    {result.rule3Mismatches.length > 5 && (
                                        <p className="text-xs text-amber-600 mt-2">ועוד {result.rule3Mismatches.length - 5} שורות... (פרוט מלא בקובץ להורדה)</p>
                                    )}
                                </div>
                            </div>
                        </div>
                    </div>
                )}
            </div>
        )}

      </main>

      {/* Config Modal */}
      {showConfig && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
                <div className="p-6 border-b border-slate-100 flex justify-between items-center sticky top-0 bg-white">
                    <h3 className="text-xl font-bold">הגדרות ספקים</h3>
                    <button onClick={() => setShowConfig(false)} className="text-slate-400 hover:text-slate-600">✕</button>
                </div>
                
                <div className="p-6 space-y-8">
                    
                    {/* Add by Name */}
                    <div className="bg-slate-50 p-5 rounded-xl border border-slate-200">
                        <h4 className="font-semibold mb-3 flex items-center gap-2">
                            <span className="w-6 h-6 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-xs">A</span>
                            הוספה לפי טקסט
                        </h4>
                        <div className="flex gap-3">
                            <input 
                                type="text" 
                                placeholder="טקסט בפרטים (לדוגמה: בזק)" 
                                className="flex-1 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
                                value={newNameKey}
                                onChange={e => setNewNameKey(e.target.value)}
                            />
                            <input 
                                type="text" 
                                placeholder="מס' ספק" 
                                className="w-32 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
                                value={newNameVal}
                                onChange={e => setNewNameVal(e.target.value)}
                            />
                            <button 
                                onClick={() => {
                                    if(newNameKey && newNameVal) {
                                        const next = {...config, nameMap: {...config.nameMap, [newNameKey]: newNameVal}};
                                        saveConfig(next);
                                        setNewNameKey(""); setNewNameVal("");
                                    }
                                }}
                                className="bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700"
                            >הוסף</button>
                        </div>
                    </div>

                    {/* Add by Amount */}
                    <div className="bg-slate-50 p-5 rounded-xl border border-slate-200">
                        <h4 className="font-semibold mb-3 flex items-center gap-2">
                            <span className="w-6 h-6 rounded-full bg-pink-100 text-pink-600 flex items-center justify-center text-xs">#</span>
                            הוספה לפי סכום מוחלט
                        </h4>
                        <div className="flex gap-3">
                            <input 
                                type="number" 
                                placeholder="סכום (לדוגמה: 150.20)" 
                                className="flex-1 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
                                value={newAmtKey}
                                onChange={e => setNewAmtKey(e.target.value)}
                            />
                            <input 
                                type="text" 
                                placeholder="מס' ספק" 
                                className="w-32 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none"
                                value={newAmtVal}
                                onChange={e => setNewAmtVal(e.target.value)}
                            />
                            <button 
                                onClick={() => {
                                    if(newAmtKey && newAmtVal) {
                                        const k = Math.abs(parseFloat(newAmtKey)).toFixed(2);
                                        const next = {...config, amountMap: {...config.amountMap, [k]: newAmtVal}};
                                        saveConfig(next);
                                        setNewAmtKey(""); setNewAmtVal("");
                                    }
                                }}
                                className="bg-pink-600 text-white px-4 py-2 rounded-lg hover:bg-pink-700"
                            >הוסף</button>
                        </div>
                    </div>

                    {/* Import Excel */}
                    <div className="pt-4 border-t border-slate-100">
                        <div className="flex justify-between items-center mb-3">
                             <h4 className="font-semibold">יבוא מקובץ אקסל</h4>
                             <button 
                                onClick={downloadTemplate}
                                className="text-emerald-600 hover:text-emerald-700 text-sm font-medium flex items-center gap-1 transition-colors"
                                title="הורד קובץ תבנית למילוי"
                            >
                                <FileSpreadsheet size={16} />
                                הורד תבנית
                            </button>
                        </div>
                        <p className="text-sm text-slate-500 mb-3">ניתן לטעון רשימה המכילה עמודות "פרטים", "סכום" ו-"מס' ספק".</p>
                        <label className="flex items-center justify-center gap-2 cursor-pointer bg-white border border-slate-300 text-slate-700 px-4 py-2 rounded-lg hover:bg-slate-50 transition-colors w-full">
                            <Upload size={16} />
                            <span>בחר קובץ אקסל...</span>
                            <input type="file" accept=".xlsx,.xls" onChange={importConfigExcel} className="hidden" />
                        </label>
                    </div>

                </div>
            </div>
        </div>
      )}
    </div>
  );
}

export default App;
