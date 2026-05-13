/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Upload, 
  Copy, 
  Check, 
  Trash2, 
  FileSpreadsheet, 
  MessageSquare,
  Clock,
  ExternalLink,
  Search,
  Download,
  AlertCircle,
  LayoutDashboard,
  History
} from 'lucide-react';
import { cn } from './lib/utils';

interface StoreData {
  toko: string;
  namaToko?: string;
  nilai: string | number;
}

const GREETINGS = [
  { label: 'Pagi', value: 'pagi' },
  { label: 'Siang', value: 'siang' },
  { label: 'Sore', value: 'sore' },
];

export default function App() {
  const [data, setData] = useState<StoreData[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [greeting, setGreeting] = useState<string>('siang');
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Initialize greeting based on current time
  React.useEffect(() => {
    const hour = new Date().getHours();
    if (hour >= 4 && hour < 11) setGreeting('pagi');
    else if (hour >= 11 && hour < 17) setGreeting('siang');
    else setGreeting('sore');
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const dataBuffer = event.target?.result;
        if (!dataBuffer) return;

        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (jsonData.length < 2) {
          alert("File Excel kosong atau tidak memiliki data.");
          return;
        }

        // Find headers
        const headers = jsonData[0].map(h => String(h).trim().toUpperCase());
        const tokoIdx = headers.indexOf('TOKO');
        const namaTokoIdx = headers.indexOf('NAMA TOKO');
        const nilaiIdx = headers.indexOf('NILAI');

        if (tokoIdx === -1 || nilaiIdx === -1) {
          alert("Gagal menemukan kolom 'TOKO' atau 'NILAI'. Pastikan nama kolom di baris pertama sudah benar.");
          return;
        }

        const parsedData = (jsonData.slice(1).map((row) => {
          const tokoValue = row[tokoIdx];
          const nilaiValue = row[nilaiIdx];
          
          if (!tokoValue) return null;

          return {
            toko: String(tokoValue).trim(),
            namaToko: namaTokoIdx !== -1 ? String(row[namaTokoIdx] || '').trim() : '',
            nilai: nilaiValue !== undefined ? nilaiValue : 0,
          };
        }).filter(item => item !== null) as StoreData[]);

        if (parsedData.length === 0) {
          alert("Tidak ada data valid yang ditemukan di file Excel tersebut.");
          setFileName(null);
          return;
        }

        setData(parsedData);
      } catch (error) {
        console.error("Error parsing Excel:", error);
        alert("Terjadi kesalahan saat membaca file. Silakan pastikan file adalah Excel yang valid (.xlsx atau .xls)");
        setFileName(null);
      }
    };
    
    reader.onerror = () => {
      alert("Gagal membaca file dari perangkat.");
      setFileName(null);
    };

    reader.readAsArrayBuffer(file);
  };

  const getMessage = (item: StoreData) => {
    const formattedNilai = typeof item.nilai === 'number' 
      ? new Intl.NumberFormat('id-ID').format(item.nilai)
      : item.nilai;

    return `Selamat ${greeting} bapak ibu ${item.toko}, berikut kurset setoran hari ini senilai Rp. ${formattedNilai}, mohon bisa segera dilunasi. Jika ada potong sales bisa dilaporkan melalui link berikut bit.ly/LaporanSetorSalesJombang. Terimakasih.`;
  };

  const copyToClipboard = (text: string, index: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIndex(index);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const filteredData = data.filter(item => 
    item.toko.toString().toLowerCase().includes(searchTerm.toLowerCase()) || 
    (item.namaToko?.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  const resetData = () => {
    setData([]);
    setFileName(null);
    setSearchTerm('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const copyAll = () => {
    const allMessages = data.map(item => getMessage(item)).join('\n\n---\n\n');
    navigator.clipboard.writeText(allMessages);
    alert('Semua pesan telah disalin ke clipboard!');
  };

  return (
    <div className="h-screen bg-slate-50 font-sans flex flex-col overflow-hidden text-slate-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-8 py-4 flex justify-between items-center shrink-0">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 w-8 h-8 rounded-lg flex items-center justify-center shadow-lg shadow-blue-100">
            <FileSpreadsheet size={18} className="text-white" />
          </div>
          <h1 className="text-xl font-bold tracking-tight text-slate-800">
            Kurset Converter <span className="text-blue-600 text-sm font-medium ml-1">Jombang Region</span>
          </h1>
        </div>
        <div className="flex items-center gap-4">
          <div className="hidden sm:flex bg-slate-100 p-1 rounded-full">
            <button className="px-4 py-1.5 text-xs font-semibold bg-white shadow-sm rounded-full flex items-center gap-2">
              <LayoutDashboard size={14} /> Convert Mode
            </button>
            <button className="px-4 py-1.5 text-xs font-medium text-slate-500 flex items-center gap-2">
              <History size={14} /> History
            </button>
          </div>
          <div className="w-8 h-8 rounded-full bg-slate-200 border border-slate-300 overflow-hidden ring-2 ring-slate-100 ring-offset-2">
            <img src="https://api.dicebear.com/7.x/avataaars/svg?seed=retail" alt="User" className="w-full h-full object-cover" />
          </div>
        </div>
      </header>

      {data.length === 0 ? (
        <main className="flex-1 flex items-center justify-center p-6">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="max-w-xl w-full text-center p-12 bg-white rounded-[2rem] border border-slate-200 shadow-2xl shadow-slate-200/50"
          >
            <div className="mb-8 inline-flex p-6 bg-blue-50 text-blue-600 rounded-3xl ring-8 ring-blue-50/50">
              <Upload size={48} strokeWidth={1.5} />
            </div>
            <h2 className="text-3xl font-black mb-4 tracking-tight text-slate-800">Mulai Konversi Data</h2>
            <p className="text-slate-500 mb-10 max-w-sm mx-auto leading-relaxed">
              Generate ratusan pesan sapaan toko dari file Excel Anda dalam hitungan detik.
            </p>
            
            <label className="relative group cursor-pointer inline-block w-full">
              <input 
                type="file" 
                ref={fileInputRef}
                accept=".xlsx, .xls" 
                className="hidden" 
                onChange={handleFileUpload}
              />
              <div className="bg-blue-600 text-white w-full py-5 rounded-2xl font-bold shadow-xl shadow-blue-200 hover:bg-blue-700 hover:-translate-y-1 transition-all flex items-center justify-center gap-3 active:scale-95 text-lg">
                <Download size={24} />
                Pilih File Excel
              </div>
            </label>
            
            <div className="mt-12 flex justify-center gap-8">
              <div className="flex flex-col items-center gap-1 group">
                <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center font-bold text-slate-400 border border-slate-100 group-hover:text-blue-600 group-hover:border-blue-200 transition-colors">T</div>
                <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Toko</span>
              </div>
              <div className="flex flex-col items-center gap-1 group">
                <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center font-bold text-slate-400 border border-slate-100 group-hover:text-blue-600 group-hover:border-blue-200 transition-colors">N</div>
                <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">Nilai</span>
              </div>
            </div>
          </motion.div>
        </main>
      ) : (
        <main className="flex-1 flex gap-6 p-6 overflow-hidden">
          {/* Sidebar: Controls */}
          <aside className="w-80 flex flex-col gap-6 shrink-0 h-full">
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
              <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">1. Data Source</h2>
              <div className="border-2 border-dashed border-blue-100 rounded-xl bg-blue-50/50 p-6 flex flex-col items-center justify-center text-center">
                <FileSpreadsheet size={32} className="text-blue-400 mb-3" />
                <p className="text-sm font-bold text-slate-700 truncate w-full">{fileName}</p>
                <p className="text-[10px] text-slate-500 mt-1 uppercase font-semibold">{data.length} Baris Terdeteksi</p>
                <button 
                  onClick={resetData}
                  className="mt-4 text-xs font-bold text-red-500 hover:bg-red-50 px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1"
                >
                  <Trash2 size={12} /> Ganti File
                </button>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex-1 flex flex-col overflow-hidden">
              <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">2. Configuration</h2>
              <div className="space-y-6 flex-1 overflow-y-auto pr-1">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-tight mb-3">Waktu Sapaan</label>
                  <div className="grid grid-cols-3 gap-2">
                    {GREETINGS.map(g => (
                      <button
                        key={g.value}
                        onClick={() => setGreeting(g.value)}
                        className={cn(
                          "py-2 text-[11px] font-bold rounded-lg transition-all",
                          greeting === g.value 
                            ? "bg-slate-900 text-white shadow-lg shadow-slate-200" 
                            : "border border-slate-200 text-slate-500 hover:bg-slate-50"
                        )}
                      >
                        {g.label}
                      </button>
                    ))}
                  </div>
                </div>
                
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-tight mb-3">Cari Toko</label>
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                    <input 
                      type="text"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      placeholder="Kode/Nama Toko..."
                      className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-100 outline-none"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-tight mb-3">Preview Template</label>
                  <div className="bg-slate-50 p-4 rounded-xl text-[11px] text-slate-500 italic border border-slate-100 leading-relaxed font-medium">
                    "Selamat [{greeting}] bapak ibu [Toko], berikut kurset setoran hari ini senilai Rp. [Nilai], mohon bisa segera dilunasi..."
                  </div>
                </div>
              </div>
              
              <div className="pt-4 border-t border-slate-100 mt-4 h-24">
                <a 
                  href="https://bit.ly/LaporanSetorSalesJombang" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="group flex items-center justify-between p-3 bg-blue-50 rounded-xl hover:bg-blue-100 transition-all border border-blue-100"
                >
                  <div className="flex items-center gap-2">
                    <div className="p-1.5 bg-blue-600 rounded-lg text-white">
                      <ExternalLink size={14} />
                    </div>
                    <span className="text-xs font-bold text-blue-700">Link Laporan</span>
                  </div>
                  <AlertCircle size={14} className="text-blue-400 group-hover:translate-x-0.5 transition-transform" />
                </a>
              </div>
            </div>
          </aside>

          {/* Main Content: Results */}
          <section className="flex-1 flex flex-col gap-6 overflow-hidden h-full">
            {/* Data Feed Table (Preview) */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-1/3 shrink-0">
              <div className="px-6 py-4 border-b border-slate-100 flex justify-between items-center">
                <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Database size={16} className="text-slate-400" /> Data Preview
                </h2>
                <span className="px-2.5 py-1 bg-green-50 text-green-600 text-[10px] font-black rounded-lg uppercase tracking-wider">Verified</span>
              </div>
              <div className="overflow-auto custom-scrollbar">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-white z-10">
                    <tr className="bg-slate-50 border-b border-slate-100">
                      <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Kode Toko</th>
                      <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Nama Toko</th>
                      <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest text-right">Nilai Setoran (IDR)</th>
                    </tr>
                  </thead>
                  <tbody className="text-[11px] text-slate-600 divide-y divide-slate-50">
                    {data.slice(0, 20).map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                        <td className="px-6 py-3 font-mono font-bold text-blue-600">{item.toko}</td>
                        <td className="px-6 py-3 font-medium">{item.namaToko || '-'}</td>
                        <td className="px-6 py-3 text-right font-bold text-slate-900">
                          {typeof item.nilai === 'number' ? new Intl.NumberFormat('id-ID').format(item.nilai) : item.nilai}
                        </td>
                      </tr>
                    ))}
                    {data.length > 20 && (
                      <tr>
                        <td colSpan={3} className="px-6 py-2 text-center text-[10px] text-slate-400 bg-slate-50 font-bold uppercase italic">
                          Showing first 20 of {data.length} rows...
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Generated Templates */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm flex-1 flex flex-col overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-100 flex justify-between items-center shrink-0">
                <div className="flex items-center gap-2">
                  <h2 className="text-sm font-bold text-slate-800">Generated Templates</h2>
                  <span className="bg-blue-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">{filteredData.length}</span>
                </div>
                <button 
                  onClick={copyAll}
                  className="text-xs font-bold text-blue-600 flex items-center gap-2 bg-blue-50 px-3 py-1.5 rounded-lg hover:bg-blue-100 transition-all border border-blue-100"
                >
                  <Copy size={14} /> Copy All
                </button>
              </div>
              <div className="p-6 overflow-y-auto custom-scrollbar flex-1 space-y-4 bg-slate-50/50">
                <AnimatePresence mode='popLayout'>
                  {filteredData.map((item, index) => (
                    <motion.div
                      key={`${item.toko}-${index}`}
                      layout
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95 }}
                      className="bg-white p-5 rounded-xl border border-slate-200 relative group hover:shadow-lg hover:shadow-slate-200/50 transition-all overflow-hidden"
                    >
                      <div className="flex justify-between items-center mb-4">
                        <div className="flex items-center gap-2">
                          <span className="text-[10px] font-black text-white bg-slate-900 px-2 py-1 rounded tracking-tighter uppercase">{item.toko}</span>
                          <span className="text-xs font-bold text-slate-400">#{index + 1}</span>
                        </div>
                        <button
                          onClick={() => copyToClipboard(getMessage(item), index)}
                          className={cn(
                            "px-4 py-2 text-[10px] font-bold uppercase tracking-widest rounded-lg flex items-center gap-2 transition-all active:scale-95",
                            copiedIndex === index 
                              ? "bg-green-100 text-green-600 shadow-sm"
                              : "bg-slate-100 text-slate-500 hover:bg-blue-600 hover:text-white"
                          )}
                        >
                          {copiedIndex === index ? <><Check size={12} /> Copied</> : <><Copy size={12} /> Copy Template</>}
                        </button>
                      </div>
                      <p className="text-sm leading-relaxed text-slate-700 font-medium">
                        Selamat {greeting} bapak ibu <span className="font-bold text-blue-600">{item.toko}</span>, berikut kurset setoran hari ini senilai <span className="font-bold text-slate-900">Rp. {typeof item.nilai === 'number' ? new Intl.NumberFormat('id-ID').format(item.nilai) : item.nilai}</span>, mohon bisa segera dilunasi. Jika ada potong sales bisa dilaporkan melalui link berikut <span className="text-blue-500 underline underline-offset-4 decoration-1 decoration-blue-200">bit.ly/LaporanSetorSalesJombang</span>. Terimakasih.
                      </p>
                    </motion.div>
                  ))}
                </AnimatePresence>
                
                {filteredData.length === 0 && (
                  <div className="text-center py-20 bg-white/50 rounded-2xl border-2 border-dashed border-slate-200">
                    <MessageSquare size={32} className="mx-auto text-slate-300 mb-2" />
                    <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">No matches found</p>
                  </div>
                )}
              </div>
            </div>
          </section>
        </main>
      )}

      {/* Global CSS for scrollbars */}
      <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 5px;
          height: 5px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #e2e8f0;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #cbd5e1;
        }
      `}</style>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 px-8 py-3 flex justify-between items-center text-[11px] text-slate-400 tracking-tight shrink-0">
        <p>© 2024 Jombang Ops • Retail Management System</p>
        <div className="flex gap-4 font-bold uppercase text-[9px] tracking-widest">
          <span className="flex items-center gap-1.5"><span className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse" /> System Online</span>
          <span>Excel/XLSX Engine</span>
        </div>
      </footer>
    </div>
  );
}

// Add Database icon from lucide (missing from imports)
const Database = ({ size, className }: { size: number, className?: string }) => (
  <svg 
    xmlns="http://www.w3.org/2000/svg" 
    width={size} 
    height={size} 
    viewBox="0 0 24 24" 
    fill="none" 
    stroke="currentColor" 
    strokeWidth="2" 
    strokeLinecap="round" 
    strokeLinejoin="round" 
    className={className}
  >
    <ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5V19A9 3 0 0 0 21 19V5"/><path d="M3 12A9 3 0 0 0 21 12"/>
  </svg>
);
