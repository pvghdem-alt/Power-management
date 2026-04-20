import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  collection, 
  onSnapshot, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  query, 
  orderBy, 
  getDocFromServer 
} from 'firebase/firestore';
import { auth, db } from './lib/firebase';
import { Reading, BLD_CONFIG, METERS } from './types';
import { 
  LayoutDashboard, 
  Table as TableIcon, 
  BarChart3, 
  LogOut, 
  LogIn, 
  Save, 
  Plus, 
  Download, 
  TrendingUp, 
  TrendingDown,
  FileText,
  ChevronRight
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer, 
  LineChart, 
  Line,
  Cell,
  LabelList
} from 'recharts';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import pptxgen from "pptxgenjs";

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Components ---

const Button = React.forwardRef<HTMLButtonElement, React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'primary' | 'secondary' | 'danger' | 'ghost' | 'success' | 'outline' }>(
  ({ className, variant = 'primary', ...props }, ref) => {
    const variants = {
      primary: 'bg-slate-900 text-white hover:bg-slate-800 shadow-sm',
      secondary: 'bg-white text-slate-900 border border-slate-200 hover:bg-slate-50 shadow-sm',
      danger: 'bg-red-600 text-white hover:bg-red-700 shadow-sm',
      success: 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-sm',
      ghost: 'hover:bg-slate-100 text-slate-600',
      outline: 'border-2 border-slate-900 text-slate-900 hover:bg-slate-900 hover:text-white',
    };
    return (
      <button
        ref={ref}
        className={cn(
          'inline-flex items-center justify-center rounded-lg px-4 py-2 text-sm font-semibold transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 disabled:pointer-events-none disabled:opacity-50',
          variants[variant],
          className
        )}
        {...props}
      />
    );
  }
);

const Card = ({ children, className, noPadding = false }: { children: React.ReactNode; className?: string; noPadding?: boolean }) => (
  <div className={cn('bg-white rounded-2xl border border-slate-200-60 shadow-sm overflow-hidden', className)}>
    <div className={cn(noPadding ? '' : 'p-6')}>
      {children}
    </div>
  </div>
);

// --- Main App ---

export default function App() {
  const [readings, setReadings] = useState<Reading[]>([]);
  const [activeTab, setActiveTab] = useState<'entry' | 'matrix' | 'analysis'>('analysis');
  const [targetMonth, setTargetMonth] = useState('');

  // Analysis state
  const [analysisMode, setAnalysisMode] = useState<'single' | 'compare' | 'growth' | 'integrated' | 'diff'>('integrated');
  const [selectedBuilding, setSelectedBuilding] = useState(BLD_CONFIG[0].v);
  const [startMonth, setStartMonth] = useState('');
  const [endMonth, setEndMonth] = useState('');
  const [baseMonth, setBaseMonth] = useState('');

  const reportRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const q = query(collection(db, 'readings'), orderBy('ts', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Reading));
      setReadings(data);
      
      // Initialize dates if not set
      const dates = (Array.from(new Set(data.map(r => r.date))) as string[]).sort((a, b) => {
        const da = parseDate(a);
        const db = parseDate(b);
        return da.y !== db.y ? da.y - db.y : da.m - db.m;
      });

      if (dates.length > 0) {
        if (!startMonth) setStartMonth(dates[Math.max(0, dates.length - 6)]);
        if (!endMonth) setEndMonth(dates[dates.length - 1]);
        if (!baseMonth) setBaseMonth(dates[dates.length - 1]);
        if (!targetMonth) setTargetMonth(getRelativeMonth(dates[dates.length - 1], 1));
      } else {
        if (!targetMonth) setTargetMonth('115年3月');
      }
    }, (error) => {
      console.error("Firestore error:", error);
    });

    return unsubscribe;
  }, []);

  // Test connection
  useEffect(() => {
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if (error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    };
    testConnection();
  }, []);

  const parseDate = (s: string) => {
    const m = s.match(/(\d+)年(\d+)月/);
    return m ? { y: parseInt(m[1]), m: parseInt(m[2]) } : { y: 0, m: 0 };
  };

  const getRelativeMonth = (s: string, o: number) => {
    const d = parseDate(s);
    if (d.y === 0) return s;
    let t = d.y * 12 + (d.m - 1) + o;
    return `${Math.floor(t / 12)}年${(t % 12) + 1}月`;
  };

  const sortedDates = useMemo(() => {
    return (Array.from(new Set(readings.map(r => r.date))) as string[]).sort((a, b) => {
      const da = parseDate(a);
      const db = parseDate(b);
      return da.y !== db.y ? da.y - db.y : da.m - db.m;
    });
  }, [readings]);

  const currentSelectedDates = useMemo(() => {
    if (!startMonth || !endMonth) return sortedDates.slice(-6);
    const startIdx = sortedDates.indexOf(startMonth);
    const endIdx = sortedDates.indexOf(endMonth);
    if (startIdx === -1 || endIdx === -1) return sortedDates.slice(-6);
    return sortedDates.slice(startIdx, endIdx + 1);
  }, [sortedDates, startMonth, endMonth]);

  const getBuildingUsage = (key: string, dates: string[]) => {
    const calcComplex = (main: string, subs: string[], dts: string[]) => {
      return dts.map(d => {
        const mainReading = readings.find(r => r.meter === main && r.date === d);
        let val = (mainReading?.usage || 0) + (mainReading?.adjustment || 0);
        subs.forEach(s => {
          const subReading = readings.find(r => r.meter === s && r.date === d);
          val -= ((subReading?.usage || 0) + (subReading?.adjustment || 0));
        });
        return Math.max(0, val);
      });
    };

    if (key === "CALC_急診") return calcComplex('01.急診大樓─總盤', ['02.急診大樓─後半段'], dates);
    if (key === "CALC_AB") return calcComplex('14.AB棟', [
      '07.役男宿舍', '08.七病房', '17.懷遠堂', '09.動力中心', '10.水塔', 
      '11.汙水處理廠', '12.廚房', '13.精神科大樓', '15.松柏園', '16.廢棄物處理廠'
    ], dates);

    return dates.map(d => {
      const r = readings.find(rd => rd.meter === key && rd.date === d);
      return (r?.usage || 0) + (r?.adjustment || 0);
    });
  };

  const handleSaveBatch = async (e: React.FormEvent) => {
    e.preventDefault();
    const form = e.target as HTMLFormElement;
    const inputs = form.querySelectorAll('input[data-meter]');
    
    try {
      for (const input of Array.from(inputs) as HTMLInputElement[]) {
        const meter = input.dataset.meter!;
        const value = parseFloat(input.value);
        const prevValue = parseFloat(input.dataset.prev || '0');
        const adjustment = parseFloat((form.querySelector(`input[data-adj="${meter}"]`) as HTMLInputElement).value || '0');
        
        if (!isNaN(value)) {
          await addDoc(collection(db, 'readings'), {
            date: targetMonth,
            meter,
            value,
            adjustment,
            usage: value - prevValue,
            ts: Date.now()
          });
        }
      }
      alert('儲存成功');
      form.reset();
    } catch (error) {
      console.error("Save failed:", error);
      alert('儲存失敗');
    }
  };

  const handleMatrixSave = async (id: string, field: string, value: string | number) => {
    try {
      if (id === 'new') {
        // Handle new entry logic if needed
      } else {
        const docRef = doc(db, 'readings', id);
        await updateDoc(docRef, { [field]: Number(value) });
        
        // Recalculate usage for this meter
        const meterName = readings.find(r => r.id === id)?.meter;
        if (meterName) {
          const meterReadings = readings
            .filter(r => r.meter === meterName)
            .sort((a, b) => {
              const da = parseDate(a.date);
              const db = parseDate(b.date);
              return da.y !== db.y ? da.y - db.y : da.m - db.m;
            });
          
          for (let i = 0; i < meterReadings.length; i++) {
            const current = meterReadings[i];
            const prev = meterReadings[i-1];
            const usage = prev ? current.value - prev.value : 0;
            if (current.usage !== usage) {
              await updateDoc(doc(db, 'readings', current.id!), { usage });
            }
          }
        }
      }
    } catch (error) {
      console.error("Matrix update failed:", error);
    }
  };

  const exportPDF = async () => {
    if (!reportRef.current) return;
    
    // Add a temporary class to ensure full rendering if needed
    const element = reportRef.current;
    const canvas = await html2canvas(element, { 
      scale: 2,
      useCORS: true,
      logging: false,
      windowWidth: element.scrollWidth,
      windowHeight: element.scrollHeight
    });
    
    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();
    
    const imgProps = pdf.getImageProperties(imgData);
    const imgWidth = pdfWidth;
    const imgHeight = (imgProps.height * imgWidth) / imgProps.width;
    
    let heightLeft = imgHeight;
    let position = 0;

    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
    heightLeft -= pdfHeight;

    while (heightLeft >= 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= pdfHeight;
    }
    
    pdf.save(`電力管理報告_${new Date().toLocaleDateString()}.pdf`);
  };

  const exportPPT = async () => {
    if (!reportRef.current) return;
    
    // Create PPTX instance
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    // Slide 1: Cover
    let slide = pres.addSlide();
    slide.background = { fill: 'FFFFFF' };
    
    slide.addText("電力管理分析報告", {
      x: 0, y: '35%', w: '100%', h: 1,
      fontSize: 44, bold: true, color: '0f172a',
      align: 'center', fontFace: 'Inter'
    });
    
    slide.addText(`龍泉分院各大樓電力度數 | ${baseMonth}`, {
      x: 0, y: '50%', w: '100%', h: 0.5,
      fontSize: 18, color: '64748b',
      align: 'center', fontFace: 'Inter'
    });

    const captureToPage = async (id: string, title?: string) => {
      const el = document.getElementById(id);
      if (!el) return;
      
      const canvas = await html2canvas(el, { 
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });
      const data = canvas.toDataURL('image/png');
      
      let s = pres.addSlide();
      if (title) {
        s.addText(title, { x: 0.5, y: 0.2, w: '90%', fontSize: 18, bold: true, color: '0f172a' });
      }
      
      // Calculate aspect ratio to prevent stretching
      const imgWidth = 9; // Max width in inches for 16x9 (approx)
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      
      // If height is too much, scale down the width
      let finalW = imgWidth;
      let finalH = imgHeight;
      if (finalH > 4.5) { // Leave room for title and padding
        finalH = 4.5;
        finalW = (canvas.width * finalH) / canvas.height;
      }

      s.addImage({
        data,
        x: (10 - finalW) / 2, // Center horizontally
        y: title ? 1.0 : 0.5,
        w: finalW,
        h: finalH
      });
    };

    // Capture individual components
    await captureToPage('ppt-section-0-chart', '各大樓用電度數比較分析 (六個月趨勢)');
    await captureToPage('ppt-section-1-chart', '各大樓用電度數比較與增減情形 (對照圖彙整)');
    await captureToPage('ppt-page-2', '各大樓電力度數增減分析表');
    await captureToPage('ppt-usage-matrix', '各大樓各月份電力度數一覽表');
    
    // Multi-page trend analysis
    await captureToPage('ppt-trend-1', '全院趨勢分析 (1/3)');
    await captureToPage('ppt-trend-2', '全院趨勢分析 (2/3)');
    await captureToPage('ppt-trend-3', '全院趨勢分析 (3/3)');

    pres.writeFile({ fileName: `電力管理報告_${new Date().toLocaleDateString()}.pptx` });
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Header */}
      <header className="bg-white backdrop-blur-md border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-slate-900 p-2.5 rounded-xl shadow-lg">
              <LayoutDashboard className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-xl tracking-tight text-slate-900">電力管理系統</h1>
              <p className="text-xs font-medium text-slate-500 uppercase tracking-widest">龍泉分院數位平台</p>
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <nav className="flex items-center gap-1 sm:gap-2 bg-slate-100-50 p-1.5 rounded-xl border border-slate-200-50">
              <Button 
                variant={activeTab === 'entry' ? 'primary' : 'ghost'} 
                onClick={() => setActiveTab('entry')}
                className={cn("gap-2 px-5 py-2.5 rounded-lg transition-all", activeTab === 'entry' && "shadow-md")}
              >
                <Plus className="w-4 h-4" />
                <span className="hidden md:inline">每月抄表</span>
              </Button>
              <Button 
                variant={activeTab === 'matrix' ? 'primary' : 'ghost'} 
                onClick={() => setActiveTab('matrix')}
                className={cn("gap-2 px-5 py-2.5 rounded-lg transition-all", activeTab === 'matrix' && "shadow-md")}
              >
                <TableIcon className="w-4 h-4" />
                <span className="hidden md:inline">矩陣編輯</span>
              </Button>
              <Button 
                variant={activeTab === 'analysis' ? 'primary' : 'ghost'} 
                onClick={() => setActiveTab('analysis')}
                className={cn("gap-2 px-5 py-2.5 rounded-lg transition-all", activeTab === 'analysis' && "shadow-md")}
              >
                <BarChart3 className="w-4 h-4" />
                <span className="hidden md:inline">趨勢分析</span>
              </Button>
            </nav>

            <div className="h-10 w-px bg-slate-200 hidden lg:block mx-2" />

            <div className="hidden lg:flex items-center gap-3 text-sm font-bold text-slate-600 bg-slate-50 px-4 py-2 rounded-xl border border-slate-200">
              <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                  <span className="uppercase tracking-wider text-[10px] text-slate-400">分析範圍</span>
                  <select 
                    value={startMonth} 
                    onChange={(e) => setStartMonth(e.target.value)}
                    className="bg-transparent outline-none cursor-pointer text-slate-900 font-bold"
                  >
                    {sortedDates.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                  <ChevronRight className="w-3 h-3 text-slate-300" />
                  <select 
                    value={endMonth} 
                    onChange={(e) => setEndMonth(e.target.value)}
                    className="bg-transparent outline-none cursor-pointer text-slate-900 font-bold"
                  >
                    {sortedDates.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
                
                <div className="w-px h-4 bg-slate-200" />
                
                <div className="flex items-center gap-2">
                  <span className="uppercase tracking-wider text-[10px] text-slate-400">基準月份</span>
                  <select 
                    value={baseMonth} 
                    onChange={(e) => setBaseMonth(e.target.value)}
                    className="bg-transparent outline-none cursor-pointer text-slate-900 font-bold"
                  >
                    {sortedDates.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl mx-auto w-full p-4 sm:p-6 space-y-6">
        {activeTab === 'entry' && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <Card className="p-0 overflow-hidden">
              <form onSubmit={handleSaveBatch}>
                <div className="p-8 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-6 bg-white">
                  <div className="space-y-1">
                    <h2 className="text-xl font-bold text-slate-900">每月抄表作業</h2>
                    <p className="text-sm text-slate-500">請輸入各電表本月讀數，系統將自動計算用電量</p>
                  </div>
                  <div className="flex items-center gap-4">
                    <div className="flex items-center gap-3 bg-slate-50 px-4 py-2 rounded-xl border border-slate-200">
                      <label className="text-sm font-bold text-slate-600">抄表月份</label>
                      <input 
                        type="text" 
                        value={targetMonth}
                        onChange={(e) => setTargetMonth(e.target.value)}
                        className="bg-transparent font-bold text-slate-900 outline-none w-28 text-center"
                        placeholder="115年3月"
                      />
                    </div>
                    <Button type="submit" variant="success" className="gap-2 px-6 h-11">
                      <Save className="w-4 h-4" />
                      儲存當月數據
                    </Button>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50-50 border-b border-slate-200-60">
                        <th className="px-8 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider">電表名稱</th>
                        <th className="px-8 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider">上月讀數</th>
                        <th className="px-8 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider">本月讀數</th>
                        <th className="px-8 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider">修正度數</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {METERS.map(meter => {
                        const lastReading = readings
                          .filter(r => r.meter === meter)
                          .sort((a, b) => b.ts - a.ts)[0];
                        return (
                          <tr key={meter} className="hover:bg-slate-50-50 transition-colors group">
                            <td className="px-8 py-5">
                              <div className="font-bold text-slate-900 group-hover:text-blue-600 transition-colors">{meter}</div>
                            </td>
                            <td className="px-8 py-5">
                              <span className="font-mono text-slate-400">{(lastReading?.value || 0).toLocaleString()}</span>
                            </td>
                            <td className="px-8 py-5">
                              <input 
                                type="number" 
                                data-meter={meter}
                                data-prev={lastReading?.value || 0}
                                className="w-full max-w-[160px] bg-slate-50 border border-slate-200 rounded-lg px-4 py-2 font-mono font-bold text-slate-900 focus:bg-white focus:ring-2 focus:ring-slate-900 focus:border-transparent outline-none transition-all"
                                placeholder="輸入讀數..."
                              />
                            </td>
                            <td className="px-8 py-5">
                              <input 
                                type="number" 
                                data-adj={meter}
                                defaultValue="0"
                                className="w-full max-w-[120px] bg-amber-50-50 border border-amber-100 rounded-lg px-4 py-2 font-mono font-bold text-amber-900 focus:bg-white focus:ring-2 focus:ring-amber-400 focus:border-transparent outline-none transition-all"
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </form>
            </Card>
          </div>
        )}

        {activeTab === 'matrix' && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <Card className="p-0 overflow-hidden">
              <div className="p-8 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-center bg-white gap-4">
                <div className="space-y-1">
                  <h2 className="text-xl font-bold text-slate-900">歷史數據矩陣</h2>
                  <p className="text-sm text-slate-500">點擊數值即可直接修改，系統將自動重新計算用電度數</p>
                </div>
                <div className="flex items-center gap-2 bg-blue-50 text-blue-700 px-4 py-2 rounded-full text-xs font-bold uppercase tracking-wider">
                  <div className="w-2 h-2 bg-blue-600 rounded-full animate-pulse" />
                  即時同步模式
                </div>
              </div>
              <div className="overflow-x-auto custom-scrollbar">
                <table className="w-full text-left border-collapse min-w-[1200px]">
                  <thead>
                    <tr className="bg-slate-50-50 border-b border-slate-200-60">
                      <th className="px-6 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider sticky left-0 bg-slate-50 z-20 border-r border-slate-200-60 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.05)]">電表 \ 月份</th>
                      {sortedDates.filter(d => {
                        const idx = sortedDates.indexOf(d);
                        const startIdx = sortedDates.indexOf(startMonth);
                        const endIdx = sortedDates.indexOf(endMonth);
                        return idx >= startIdx && idx <= endIdx;
                      }).map(date => (
                        <th key={date} className="px-6 py-4 font-bold text-slate-500 text-xs uppercase tracking-wider text-center min-w-[160px]">{date}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {METERS.map(meter => (
                      <tr key={meter} className="hover:bg-slate-50-30 transition-colors group">
                        <td className="px-6 py-5 font-bold text-slate-900 sticky left-0 bg-white z-20 border-r border-slate-200-60 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.05)] group-hover:text-blue-600 transition-colors">{meter}</td>
                        {sortedDates.filter(d => {
                          const idx = sortedDates.indexOf(d);
                          const startIdx = sortedDates.indexOf(startMonth);
                          const endIdx = sortedDates.indexOf(endMonth);
                          return idx >= startIdx && idx <= endIdx;
                        }).map(date => {
                          const reading = readings.find(r => r.meter === meter && r.date === date);
                          return (
                            <td key={date} className="px-6 py-5 text-center space-y-2">
                              <input 
                                type="number" 
                                defaultValue={reading?.value || ''}
                                onBlur={(e) => reading && handleMatrixSave(reading.id!, 'value', e.target.value)}
                                className="w-full text-center bg-transparent border-b-2 border-transparent hover:border-slate-200 focus:border-slate-900 font-mono font-bold text-slate-900 py-1 outline-none transition-all"
                              />
                              <div className="flex items-center justify-center gap-1.5">
                                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">校正:</span>
                                <input 
                                  type="number" 
                                  defaultValue={reading?.adjustment || 0}
                                  onBlur={(e) => reading && handleMatrixSave(reading.id!, 'adjustment', e.target.value)}
                                  className="w-16 text-center bg-amber-50 text-amber-900 font-mono text-[11px] font-bold rounded-md border border-amber-100-50 py-0.5 outline-none focus:ring-1 focus:ring-amber-400"
                                />
                              </div>
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          </div>
        )}

        {activeTab === 'analysis' && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
            {/* Analysis Controls */}
            <Card className="p-4 bg-white backdrop-blur-sm sticky top-[80px] z-20 border-slate-200-60 shadow-md">
              <div className="flex flex-wrap items-center gap-6">
                <div className="flex bg-slate-200-50 p-1 rounded-xl border border-slate-200-50">
                  {[
                    { id: 'integrated', label: '分析報告', icon: FileText },
                    { id: 'single', label: '單棟趨勢', icon: TrendingUp },
                  ].map(mode => (
                    <button
                      key={mode.id}
                      onClick={() => setAnalysisMode(mode.id as any)}
                      className={cn(
                        'flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all',
                        analysisMode === mode.id ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500 hover:text-slate-700 hover:bg-white'
                      )}
                    >
                      <mode.icon className="w-4 h-4" />
                      {mode.label}
                    </button>
                  ))}
                </div>

                <div className="h-8 w-px bg-slate-200 hidden lg:block" />

                {analysisMode === 'single' && (
                  <div className="flex items-center gap-3 text-sm font-bold text-slate-600">
                    <span className="uppercase tracking-wider text-[10px]">分析對象</span>
                    <select 
                      value={selectedBuilding} 
                      onChange={(e) => setSelectedBuilding(e.target.value)}
                      className="bg-slate-50 border border-slate-200 rounded-lg px-4 py-1.5 outline-none font-bold text-slate-900 cursor-pointer"
                    >
                      {BLD_CONFIG.map(b => <option key={b.v} value={b.v}>{b.l}</option>)}
                    </select>
                  </div>
                )}

                <div className="ml-auto flex gap-2">
                  <Button onClick={exportPPT} variant="outline" className="gap-2 px-6">
                    <FileText className="w-4 h-4" />
                    匯出 PPT 簡報
                  </Button>
                  <Button onClick={exportPDF} variant="outline" className="gap-2 px-6">
                    <Download className="w-4 h-4" />
                    匯出 PDF 報告
                  </Button>
                </div>
              </div>
            </Card>

            {/* Charts Area */}
            <div ref={reportRef} className="space-y-12 bg-white p-12 rounded-[2rem] border border-slate-200-60 shadow-xl">
              {analysisMode === 'single' && (
                <div className="space-y-16">
                  <div className="text-center space-y-3">
                    <div className="inline-block px-4 py-1.5 bg-blue-50 text-blue-700 rounded-full text-xs font-black uppercase tracking-[0.2em] mb-2">Building Analysis</div>
                    <h2 className="text-4xl font-black text-slate-900 tracking-tight">{BLD_CONFIG.find(b => b.v === selectedBuilding)?.l}</h2>
                    <p className="text-slate-400 font-medium">數據區間：{startMonth} — {endMonth}</p>
                  </div>
                  
                  <div className="grid grid-cols-1 gap-12">
                    <div className="space-y-4">
                      <h4 className="text-sm font-black text-slate-900 uppercase tracking-widest flex items-center gap-3">
                        <div className="w-8 h-1 bg-slate-900 rounded-full" />
                        用電趨勢分佈 (柱狀圖)
                      </h4>
                      <div className="h-[400px] w-full bg-slate-50-50 rounded-3xl p-8 border border-slate-100">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={
                            sortedDates.slice(sortedDates.indexOf(startMonth), sortedDates.indexOf(endMonth) + 1).map(date => {
                              const prevMonth = getRelativeMonth(date, -1);
                              const usage = getBuildingUsage(selectedBuilding, [date, prevMonth]);
                              const current = usage[0];
                              const previous = usage[1];
                              
                              return {
                                date,
                                base: Math.min(current, previous),
                                increase: current > previous ? current - previous : 0,
                                decrease: previous > current ? previous - current : 0
                              };
                            })
                          }>
                            <CartesianGrid strokeDasharray="8 8" vertical={false} stroke="#e2e8f0" />
                            <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 13, fontWeight: 700 }} dy={15} />
                            <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 13, fontWeight: 700 }} />
                            <Tooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderRadius: '12px', border: 'none', color: '#fff', padding: '12px' }}
                              itemStyle={{ color: '#fff', fontWeight: 'bold' }}
                              cursor={{ fill: '#f1f5f9', opacity: 0.4 }}
                            />
                            <Legend 
                              verticalAlign="top" 
                              align="center" 
                              iconType="rect" 
                              wrapperStyle={{ paddingBottom: '20px' }}
                              formatter={(value) => {
                                if (value === 'base') return '基礎';
                                if (value === 'increase') return '增';
                                if (value === 'decrease') return '減';
                                return value;
                              }}
                            />
                            <Bar dataKey="base" stackId="a" fill="#3b82f6" radius={[0, 0, 0, 0]} />
                            <Bar dataKey="increase" stackId="a" fill="#ef4444" radius={[4, 4, 0, 0]}>
                              <LabelList dataKey="increase" position="top" formatter={(v: number) => v > 0 ? v.toLocaleString() : ''} style={{ fill: '#ef4444', fontSize: 12, fontWeight: 900 }} />
                            </Bar>
                            <Bar dataKey="decrease" stackId="a" fill="#10b981" radius={[4, 4, 0, 0]}>
                              <LabelList dataKey="decrease" position="top" formatter={(v: number) => v > 0 ? v.toLocaleString() : ''} style={{ fill: '#10b981', fontSize: 12, fontWeight: 900 }} />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="space-y-4">
                      <h4 className="text-sm font-black text-slate-900 uppercase tracking-widest flex items-center gap-3">
                        <div className="w-8 h-1 bg-blue-600 rounded-full" />
                        用電波動分析 (折線圖)
                      </h4>
                      <div className="h-[400px] w-full bg-slate-50-50 rounded-3xl p-8 border border-slate-100">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={
                            sortedDates.slice(sortedDates.indexOf(startMonth), sortedDates.indexOf(endMonth) + 1).map(date => ({
                              date,
                              usage: getBuildingUsage(selectedBuilding, [date])[0]
                            }))
                          }>
                            <CartesianGrid strokeDasharray="8 8" vertical={false} stroke="#e2e8f0" />
                            <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 13, fontWeight: 700 }} dy={15} />
                            <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 13, fontWeight: 700 }} />
                            <Tooltip />
                            <Line type="monotone" dataKey="usage" stroke="#2563eb" strokeWidth={4} dot={{ r: 6, fill: '#2563eb', strokeWidth: 3, stroke: '#fff' }} activeDot={{ r: 10, strokeWidth: 0 }} />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {analysisMode === 'compare' && (
                <div className="space-y-12" id="ppt-page-3">
                  <div className="text-center space-y-3">
                    <div className="inline-block px-4 py-1.5 bg-emerald-50 text-emerald-700 rounded-full text-xs font-black uppercase tracking-[0.2em] mb-2">Comparative Analysis</div>
                    <h2 className="text-4xl font-black text-slate-900 tracking-tight">各大樓用電度數比較</h2>
                    <p className="text-slate-400 font-medium">數據區間：{startMonth} — {endMonth}</p>
                  </div>
                  <div className="h-[600px] w-full bg-slate-50-50 rounded-3xl p-8 border border-slate-100">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart 
                        data={BLD_CONFIG.map(b => ({
                          name: b.l,
                          ...Object.fromEntries(
                            currentSelectedDates.map(d => [d, getBuildingUsage(b.v, [d])[0]])
                          )
                        }))}
                      >
                        <CartesianGrid strokeDasharray="8 8" vertical={false} stroke="#e2e8f0" />
                        <XAxis 
                          dataKey="name" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fill: '#0f172a', fontSize: 13, fontWeight: 800 }}
                          interval={0}
                          angle={-45}
                          textAnchor="end"
                          height={80}
                        />
                        <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 13, fontWeight: 700 }} />
                        <Tooltip />
                        <Legend verticalAlign="top" align="right" iconType="circle" wrapperStyle={{ paddingBottom: '20px' }} />
                        {currentSelectedDates.map((date, idx) => (
                          <Bar key={date} dataKey={date} fill={['#0f172a', '#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#6366f1', '#ec4899', '#f97316'][idx % 10]} radius={[6, 6, 0, 0]} barSize={20}>
                            <LabelList dataKey={date} position="top" formatter={(v: number) => v > 1000 ? (v/1000).toFixed(1) + 'k' : v} style={{ fontSize: 11, fontWeight: 900 }} />
                          </Bar>
                        ))}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              {analysisMode === 'growth' && (
                <div className="space-y-12" id="ppt-page-2">
                  <div className="text-center space-y-3">
                    <div className="inline-block px-4 py-1.5 bg-amber-50 text-amber-700 rounded-full text-xs font-black uppercase tracking-[0.2em] mb-2">Growth & Efficiency</div>
                    <h2 className="text-4xl font-black text-slate-900 tracking-tight">各大樓電力度數增減分析</h2>
                    <p className="text-slate-400 font-medium">基準月份：{baseMonth}</p>
                  </div>
                  
                  <div className="overflow-hidden rounded-3xl border border-slate-200-60 shadow-sm">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-900 text-white">
                          <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">大樓名稱</th>
                          <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">本月度數</th>
                          <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">較上月增減</th>
                          <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">較去年增減</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {BLD_CONFIG.map(b => {
                          const prev = getRelativeMonth(baseMonth, -1);
                          const year = getRelativeMonth(baseMonth, -12);
                          const usage = getBuildingUsage(b.v, [baseMonth, prev, year]);
                          const cur = usage[0];
                          const pre = usage[1];
                          const lastY = usage[2];
                          
                          const diffM = cur - pre;
                          const perM = pre ? ((diffM / pre) * 100).toFixed(1) : '0';
                          const diffY = lastY ? cur - lastY : null;
                          const perY = lastY ? ((diffY! / lastY) * 100).toFixed(1) : null;

                          return (
                            <tr key={b.v} className="hover:bg-slate-50 transition-colors">
                              <td className="px-8 py-5 font-bold text-slate-900">{b.l}</td>
                              <td className="px-8 py-5 text-slate-700 font-mono font-bold">{cur.toLocaleString()}</td>
                              <td className="px-8 py-5">
                                <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffM > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                  {diffM > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                  {diffM > 0 ? '+' : ''}{diffM.toLocaleString()} ({perM}%)
                                </div>
                              </td>
                              <td className="px-8 py-5">
                                {diffY !== null ? (
                                  <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffY > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                    {diffY > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                    {diffY > 0 ? '+' : ''}{diffY.toLocaleString()} ({perY}%)
                                  </div>
                                ) : (
                                  <span className="text-slate-300 text-[10px] font-bold uppercase tracking-widest">No Data</span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                        <tr className="bg-slate-100 font-black border-t-2 border-slate-900">
                          <td className="px-8 py-5 text-slate-900 uppercase tracking-widest text-xs">全院總計</td>
                          {(() => {
                            const prev = getRelativeMonth(baseMonth, -1);
                            const year = getRelativeMonth(baseMonth, -12);
                            let totalCur = 0; let totalPre = 0; let totalLastY: number | null = 0;
                            BLD_CONFIG.forEach(b => {
                              const usage = getBuildingUsage(b.v, [baseMonth, prev, year]);
                              totalCur += usage[0]; totalPre += usage[1];
                              if (totalLastY !== null && usage[2] !== null) { totalLastY += usage[2]; } else { totalLastY = null; }
                            });
                            const diffM = totalCur - totalPre;
                            const perM = totalPre ? ((diffM / totalPre) * 100).toFixed(1) : '0';
                            const diffY = totalLastY !== null ? totalCur - totalLastY : null;
                            const perY = totalLastY !== null && totalLastY !== 0 ? ((diffY! / totalLastY) * 100).toFixed(1) : null;
                            return (
                              <>
                                <td className="px-8 py-5 text-slate-700 font-mono font-bold">{totalCur.toLocaleString()}</td>
                                <td className="px-8 py-5">
                                  <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffM > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                    {diffM > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                    {diffM > 0 ? '+' : ''}{diffM.toLocaleString()} ({perM}%)
                                  </div>
                                </td>
                                <td className="px-8 py-5">
                                  {diffY !== null ? (
                                    <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffY > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                      {diffY > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                      {diffY > 0 ? '+' : ''}{diffY.toLocaleString()} ({perY}%)
                                    </div>
                                  ) : (
                                    <span className="text-slate-300 text-[10px] font-bold uppercase tracking-widest">No Data</span>
                                  )}
                                </td>
                              </>
                            );
                          })()}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {analysisMode === 'integrated' && (
                <div className="space-y-20">
                  <div className="text-center space-y-6 border-b-4 border-slate-900 pb-12">
                    <div className="flex justify-center mb-4">
                      <div className="bg-slate-900 p-4 rounded-3xl">
                        <LayoutDashboard className="w-12 h-12 text-white" />
                      </div>
                    </div>
                    <h1 className="text-6xl font-black tracking-tighter text-slate-900 uppercase">電力管理分析報告</h1>
                    <div className="flex justify-center gap-12 text-xs font-black text-slate-400 uppercase tracking-[0.3em]">
                      <span>龍泉分院電力管理系統</span>
                      <span>基準月份：{baseMonth}</span>
                    </div>
                  </div>

              <div className="space-y-16">
                {/* Section 01: Multi-month Comparison */}
                <section className="space-y-8">
                  <div className="flex items-center gap-4">
                    <div className="text-4xl font-black text-slate-200">01</div>
                    <div className="space-y-1">
                      <h3 className="text-2xl font-black text-slate-900 tracking-tight uppercase">各大樓用電度數比較分析</h3>
                      <p className="text-xs font-bold text-slate-400">顯示範圍：{startMonth} 至 {endMonth}</p>
                    </div>
                  </div>
                  <div id="ppt-section-0-chart" className="h-[520px] w-full bg-slate-50-50 rounded-[2rem] p-10 border border-slate-100 mb-8">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart 
                        data={BLD_CONFIG.map(b => {
                          const months = currentSelectedDates;
                          const usage = getBuildingUsage(b.v, months);
                          const data: any = { name: b.l };
                          months.forEach((m, i) => data[m] = usage[i]);
                          return data;
                        })}
                        margin={{ top: 30, right: 30, left: 20, bottom: 60 }}
                      >
                        <CartesianGrid strokeDasharray="8 8" vertical={false} stroke="#e2e8f0" />
                        <XAxis 
                          dataKey="name" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fill: '#64748b', fontSize: 12, fontWeight: 800 }}
                          interval={0}
                          angle={-45}
                          textAnchor="end"
                        />
                        <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 12, fontWeight: 700 }} />
                        <Tooltip 
                          cursor={{ fill: '#f1f5f9', opacity: 0.4 }}
                          contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1)' }}
                        />
                        <Legend 
                          verticalAlign="top" 
                          align="center" 
                          iconType="circle" 
                          wrapperStyle={{ paddingBottom: '30px' }}
                        />
                        {currentSelectedDates.map((m, i) => (
                          <Bar 
                            key={m} 
                            dataKey={m} 
                            fill={['#0f172a', '#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#6366f1', '#ec4899', '#f97316'][i % 10]} 
                            radius={[4, 4, 0, 0]}
                            isAnimationActive={false}
                          />
                        ))}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </section>

              <section className="space-y-8">
                <div className="flex items-center gap-4">
                  <div className="text-4xl font-black text-slate-200">02</div>
                  <div className="space-y-1">
                    <h3 className="text-2xl font-black text-slate-900 tracking-tight uppercase">各大樓用電度數比較與增減情形</h3>
                    <p className="text-xs font-bold text-slate-400">
                      基準月份：{baseMonth} (與 {getRelativeMonth(baseMonth, -1)} 比較)
                    </p>
                  </div>
                </div>
                <div id="ppt-section-1-chart" className="h-[520px] w-full bg-slate-50-50 rounded-[2rem] p-10 border border-slate-100 mb-8">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={BLD_CONFIG.map(b => {
                        const prevMonth = getRelativeMonth(baseMonth, -1);
                        const usage = getBuildingUsage(b.v, [baseMonth, prevMonth]);
                        const current = usage[0];
                        const previous = usage[1];
                        return {
                          name: b.l,
                          base: Math.min(current, previous),
                          increase: current > previous ? current - previous : 0,
                          decrease: previous > current ? previous - current : 0,
                          total: current,
                          labelTrigger: 1 // Non-zero for trigger
                        };
                      })}
                      margin={{ top: 50, right: 30, left: 20, bottom: 60 }}
                    >
                      <CartesianGrid strokeDasharray="8 8" vertical={false} stroke="#e2e8f0" />
                      <XAxis 
                        dataKey="name" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fill: '#64748b', fontSize: 12, fontWeight: 800 }}
                        interval={0}
                        angle={-45}
                        textAnchor="end"
                      />
                      <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontSize: 12, fontWeight: 700 }} />
                      <Tooltip 
                        cursor={{ fill: '#f1f5f9', opacity: 0.4 }}
                        contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1)' }}
                        formatter={(value: number, name: string, props: any) => {
                          const total = props?.payload?.total ?? 0;
                          let label = value.toLocaleString() + ' 度';
                          if (name === 'increase') label = `+${label}`;
                          if (name === 'decrease') label = `-${label}`;
                          if (name === 'labelTrigger') return null;
                          return [label, name === 'base' ? '基礎' : name === 'increase' ? '增' : '減'];
                        }}
                      />
                      <Legend 
                        verticalAlign="top" 
                        align="center" 
                        iconType="rect" 
                        wrapperStyle={{ paddingBottom: '30px' }}
                        formatter={(value) => {
                          if (value === 'base') return '基礎';
                          if (value === 'increase') return '增';
                          if (value === 'decrease') return '減';
                          return value;
                        }}
                      />
                      <Bar dataKey="base" stackId="a" fill="#3b82f6" isAnimationActive={false} />
                      <Bar dataKey="increase" stackId="a" fill="#ef4444" radius={[6, 6, 0, 0]} isAnimationActive={false} />
                      <Bar dataKey="decrease" stackId="a" fill="#10b981" radius={[6, 6, 0, 0]} isAnimationActive={false} />
                      <Bar dataKey="labelTrigger" stackId="a" fill="red" fillOpacity={0} isAnimationActive={false} legendType="none">
                        <LabelList 
                          dataKey="labelTrigger" 
                          position="top" 
                          content={(props: any) => {
                            const { x, y, width, payload } = props;
                            if (!payload) return null;
                            const totalVal = (payload.base || 0) + (payload.increase || 0);
                            const totalStr = totalVal.toLocaleString();
                            const inc = payload.increase || 0;
                            const dec = payload.decrease || 0;
                            
                            return (
                              <text x={x + width / 2} y={y - 15} textAnchor="middle" fontSize={12} fontWeight={900}>
                                <tspan fill="#000000">{totalStr}</tspan>
                                {inc > 0 && <tspan fill="#ef4444"> (+{inc.toLocaleString()})</tspan>}
                                {dec > 0 && <tspan fill="#10b981"> (-{dec.toLocaleString()})</tspan>}
                              </text>
                            );
                          }}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </section>

                  <section className="space-y-8">
                    <div className="flex items-center gap-4">
                      <div className="text-4xl font-black text-slate-200">03</div>
                      <h3 className="text-2xl font-black text-slate-900 tracking-tight uppercase">各大樓電力度數增減分析表</h3>
                    </div>
                    <div id="ppt-page-2" className="overflow-hidden rounded-3xl border border-slate-200-60 shadow-sm">
                      <table className="w-full text-left border-collapse bg-white">
                        <thead>
                          <tr className="bg-slate-900 text-white">
                            <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">大樓名稱</th>
                            <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">本月度數</th>
                            <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">較上月增減</th>
                            <th className="px-8 py-5 font-bold text-xs uppercase tracking-widest">較去年增減</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {BLD_CONFIG.map(b => {
                            const prev = getRelativeMonth(baseMonth, -1);
                            const year = getRelativeMonth(baseMonth, -12);
                            const usage = getBuildingUsage(b.v, [baseMonth, prev, year]);
                            const cur = usage[0];
                            const pre = usage[1];
                            const lastY = usage[2];
                            const diffM = cur - pre;
                            const perM = pre ? ((diffM / pre) * 100).toFixed(1) : '0';
                            const diffY = lastY ? cur - lastY : null;
                            const perY = lastY ? ((diffY! / lastY) * 100).toFixed(1) : null;
                            return (
                              <tr key={b.v} className="hover:bg-slate-50 transition-colors">
                                <td className="px-8 py-3 font-bold text-slate-900">{b.l}</td>
                                <td className="px-8 py-3 text-slate-700 font-mono font-bold">{cur.toLocaleString()}</td>
                                <td className="px-8 py-3">
                                  <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffM > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                    {diffM > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                    {diffM > 0 ? '+' : ''}{diffM.toLocaleString()} ({perM}%)
                                  </div>
                                </td>
                                <td className="px-8 py-3">
                                  {diffY !== null ? (
                                    <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffY > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                      {diffY > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                      {diffY > 0 ? '+' : ''}{diffY.toLocaleString()} ({perY}%)
                                    </div>
                                  ) : (
                                    <span className="text-slate-300 text-[10px] font-bold uppercase tracking-widest">No Data</span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                          <tr className="bg-slate-100 font-black border-t-2 border-slate-900">
                            <td className="px-8 py-4 text-slate-900 uppercase tracking-widest text-xs">全院總計</td>
                            {(() => {
                              const prev = getRelativeMonth(baseMonth, -1);
                              const year = getRelativeMonth(baseMonth, -12);
                              let totalCur = 0; let totalPre = 0; let totalLastY: number | null = 0;
                              BLD_CONFIG.forEach(b => {
                                const usage = getBuildingUsage(b.v, [baseMonth, prev, year]);
                                totalCur += usage[0]; totalPre += usage[1];
                                if (totalLastY !== null && usage[2] !== null) { totalLastY += usage[2]; } else { totalLastY = null; }
                              });
                              const diffM = totalCur - totalPre;
                              const perM = totalPre ? ((diffM / totalPre) * 100).toFixed(1) : '0';
                              const diffY = totalLastY !== null ? totalCur - totalLastY : null;
                              const perY = totalLastY !== null && totalLastY !== 0 ? ((diffY! / totalLastY) * 100).toFixed(1) : null;
                              return (
                                <>
                                  <td className="px-8 py-4 text-slate-700 font-mono font-bold">{totalCur.toLocaleString()}</td>
                                  <td className="px-8 py-4">
                                    <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffM > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                      {diffM > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                      {diffM > 0 ? '+' : ''}{diffM.toLocaleString()} ({perM}%)
                                    </div>
                                  </td>
                                  <td className="px-8 py-4">
                                    {diffY !== null ? (
                                      <div className={cn("flex items-center gap-2 px-3 py-1 rounded-full w-fit text-xs font-black", diffY > 0 ? "bg-red-50-50 text-red-600" : "bg-emerald-50-50 text-emerald-600")}>
                                        {diffY > 0 ? <TrendingUp className="w-3 h-3" /> : <TrendingDown className="w-3 h-3" />}
                                        {diffY > 0 ? '+' : ''}{diffY.toLocaleString()} ({perY}%)
                                      </div>
                                    ) : (
                                      <span className="text-slate-300 text-[10px] font-bold uppercase tracking-widest">No Data</span>
                                    )}
                                  </td>
                                </>
                              );
                            })()}
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </section>

                  <section className="space-y-8">
                    <div className="flex items-center gap-4">
                      <div className="text-4xl font-black text-slate-200">04</div>
                      <h3 className="text-2xl font-black text-slate-900 tracking-tight uppercase">各大樓各月份電力度數一覽表</h3>
                    </div>
                    <div id="ppt-usage-matrix" className="overflow-hidden rounded-3xl border border-slate-200 shadow-sm overflow-x-auto">
                      <table className="w-full text-left border-collapse bg-white min-w-[1000px]">
                        <thead>
                          <tr className="bg-slate-900 text-white">
                            <th className="px-6 py-4 font-bold text-xs uppercase tracking-widest border-r border-slate-800">大樓名稱</th>
                            {currentSelectedDates.map(date => (
                              <th key={date} className="px-6 py-4 font-bold text-xs uppercase tracking-widest text-center">{date}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {BLD_CONFIG.map(b => (
                            <tr key={b.v} className="hover:bg-slate-50 transition-colors">
                              <td className="px-6 py-3 font-bold text-slate-900 border-r border-slate-100 bg-slate-50">{b.l}</td>
                              {currentSelectedDates.map(date => {
                                const val = getBuildingUsage(b.v, [date])[0];
                                return (
                                  <td key={date} className="px-6 py-3 text-center text-slate-700 font-mono text-sm">{val.toLocaleString()}</td>
                                );
                              })}
                            </tr>
                          ))}
                          <tr className="bg-slate-100 font-black border-t-2 border-slate-900">
                            <td className="px-6 py-4 text-slate-900 border-r border-slate-200 bg-slate-100 uppercase tracking-widest text-xs">全院總計</td>
                            {currentSelectedDates.map(date => {
                              const total = BLD_CONFIG.reduce((sum, b) => sum + getBuildingUsage(b.v, [date])[0], 0);
                              return (
                                <td key={date} className="px-6 py-4 text-center text-blue-700 font-mono text-sm">
                                  {total.toLocaleString()}
                                </td>
                              );
                            })}
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </section>

                  <section className="space-y-8">
                    <div className="flex items-center gap-4">
                      <div className="text-4xl font-black text-slate-200">05</div>
                      <h3 className="text-2xl font-black text-slate-900 tracking-tight uppercase">全院各大樓趨勢分析</h3>
                    </div>
                    <div className="space-y-12">
                      {[0, 1, 2].map(pageIdx => (
                        <div key={pageIdx} id={`ppt-trend-${pageIdx + 1}`} className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 pt-8 border-t border-slate-100 first:border-0 first:pt-0">
                          {BLD_CONFIG.slice(pageIdx * 6, (pageIdx + 1) * 6).map(b => (
                            <div key={b.v} className="p-6 border border-slate-100 rounded-[2rem] bg-slate-50-50 hover:bg-white hover:shadow-xl transition-all group">
                              <h4 className="text-xs font-black text-slate-400 mb-4 text-center uppercase tracking-widest group-hover:text-slate-900 transition-colors">{b.l}</h4>
                              <div className="h-40">
                                <ResponsiveContainer width="100%" height="100%">
                                  <LineChart data={sortedDates.slice(sortedDates.indexOf(startMonth), sortedDates.indexOf(endMonth) + 1).map(d => ({ d, v: getBuildingUsage(b.v, [d])[0] }))}>
                                    <CartesianGrid strokeDasharray="4 4" vertical={false} stroke="#e2e8f0" />
                                    <XAxis dataKey="d" axisLine={false} tickLine={false} tick={{ fontSize: 8, fill: '#94a3b8' }} />
                                    <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 8, fill: '#94a3b8' }} domain={['auto', 'auto']} />
                                    <Line type="monotone" dataKey="v" stroke="#0f172a" strokeWidth={3} dot={{ r: 3, fill: '#0f172a' }}>
                                      <LabelList dataKey="v" position="top" style={{ fontSize: 8, fontWeight: 900, fill: '#0f172a' }} formatter={(v: number) => v.toLocaleString()} />
                                    </Line>
                                  </LineChart>
                                </ResponsiveContainer>
                              </div>
                            </div>
                          ))}
                        </div>
                      ))}
                    </div>
                  </section>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </main>

      <footer className="bg-white border-t border-slate-200 py-12">
        <div className="max-w-7xl mx-auto px-4 text-center space-y-2">
          <p className="text-slate-900 font-bold tracking-tight">龍泉分院電力管理系統</p>
          <p className="text-slate-400 text-xs font-medium uppercase tracking-[0.2em]">
            © {new Date().getFullYear()} Digital Management Platform · Enterprise Edition
          </p>
        </div>
      </footer>
    </div>
  );
}
