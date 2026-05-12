import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { Package, MapPin, Inbox, Box, Plus, Minus, Building2, Warehouse as WarehouseIcon } from 'lucide-react';

/* 
Data Logic (Excel Parsing)
Parse the uploaded Excel sheet using these mappings:
1. **Bonded Inventory:** Map rows where 'StorageType' = 'Bonded' or 'CustomsStatus' is 'Pending'.
2. **Warehouse Stocks:** Map rows where 'Location' matches specific Warehouse IDs.
3. **Buffer Zone:** Map rows where 'Status' = 'Arrived' but 'FinalDestination' is null.
4. **Capacity Logic:** Compare the COUNT of rows per 'Location' against a fixed or dynamic 'MaxCapacity' column.
*/

type StorageInventoryProps = {
  data: any[];
};

type LocationData = {
  id: string;
  name: string;
  empty: number;
  full: number;
  capacity?: number;
};

type SectionData = {
  id: string;
  title: string;
  icon: React.ReactNode;
  locations: LocationData[];
};

const initialSections: SectionData[] = [
  {
    id: 'bonded',
    title: 'BONDED AREA',
    icon: <Building2 className="w-5 h-5 text-amber-500" />,
    locations: []
  },
  {
    id: 'warehouse',
    title: 'WAREHOUSE',
    icon: <WarehouseIcon className="w-5 h-5 text-blue-500" />,
    locations: []
  },
  {
    id: 'buffer',
    title: 'BUFFER',
    icon: <Inbox className="w-5 h-5 text-emerald-500" />,
    locations: []
  }
];

export const StorageInventory: React.FC<StorageInventoryProps> = ({ data }) => {
  const [sections, setSections] = useState<SectionData[]>(initialSections);
  const [addingSection, setAddingSection] = useState<string | null>(null);
  const [newEntryName, setNewEntryName] = useState<string>("");
  const [manualBufferAvgStay, setManualBufferAvgStay] = useState<string>("0.0");
  
  // Manual Inventory Aging state
  const [manualAging, setManualAging] = useState({
    '1-7d': "",
    '8-15d': "",
    '16-30d': "",
    '30d+': ""
  });

  const handleUpdate = (sectionId: string, locId: string, field: 'empty' | 'full' | 'capacity', value: number) => {
    setSections(prev => prev.map(sec => {
      if (sec.id !== sectionId) return sec;
      return {
        ...sec,
        locations: sec.locations.map(loc => {
          if (loc.id !== locId) return loc;
          const currentVal = loc[field] || 0;
          return { ...loc, [field]: Math.max(0, currentVal + value) };
        })
      };
    }));
  };

  const handeInputMatch = (sectionId: string, locId: string, field: 'empty' | 'full' | 'capacity', val: string) => {
      let num = parseInt(val, 10);
      if (isNaN(num)) num = 0;
      setSections(prev => prev.map(sec => {
        if (sec.id !== sectionId) return sec;
        return {
            ...sec,
            locations: sec.locations.map(loc => {
                if (loc.id !== locId) return loc;
                return { ...loc, [field]: Math.max(0, num) };
            })
        };
      }));
  }

  const confirmAddLocation = (sectionId: string) => {
    if (!newEntryName.trim()) return;
    setSections(prev => prev.map(sec => {
      if (sec.id !== sectionId) return sec;
      return {
        ...sec,
        locations: [...sec.locations, { id: 'loc_' + Date.now(), name: newEntryName.trim().toUpperCase(), empty: 0, full: 0, capacity: 0 }]
      };
    }));
    setAddingSection(null);
    setNewEntryName("");
  };

  const cancelAddLocation = () => {
    setAddingSection(null);
    setNewEntryName("");
  };

  const metrics = useMemo(() => {
    let bondedCount = 0;
    let bufferCount = 0;
    let warehouseCount = 0;
    let warehouseCounts: Record<string, number> = {};
    let MAX_CAPACITY = 10000; // default fallback higher since we have thousands manually

    // Calculate manual totals
    let manualBonded = 0;
    let manualWarehouse = 0;
    let manualBuffer = 0;
    let manualCapacity = 0;
    let hasManualEntries = false;
    
    sections.forEach(sec => {
      if (sec.locations.length > 0) hasManualEntries = true;
      sec.locations.forEach(loc => {
        const totalLoc = (loc.empty || 0) + (loc.full || 0);
        manualCapacity += loc.capacity || 0;
        if (sec.id === 'bonded') manualBonded += totalLoc;
        if (sec.id === 'warehouse') manualWarehouse += totalLoc;
        if (sec.id === 'buffer') manualBuffer += totalLoc;
      });
    });

    // 3. Inventory Aging
    let aging = {
      '1-7d': 0,
      '8-15d': 0,
      '16-30d': 0,
      '30d+': 0
    };

    let totalWaitTime = 0;
    
    data.forEach(row => {
      const storageType = String(row['StorageType'] || '').trim().toLowerCase();
      const customsStatus = String(row['CustomsStatus'] || '').trim().toLowerCase();
      const status = String(row['Status'] || row['STATUS'] || '').trim().toLowerCase();
      const location = String(row['Location'] || row['BONDED WAREHOUSE'] || 'Default WH').trim();
      const finalDest = row['FinalDestination'] || row['FINAL DESTINATION'];
      
      const rowMaxCap = parseInt(String(row['MaxCapacity'] || row['MAX_CAPACITY']), 10);
      if (!isNaN(rowMaxCap) && rowMaxCap > MAX_CAPACITY) {
        MAX_CAPACITY = rowMaxCap;
      }

      // Bonded
      if (storageType === 'bonded' || customsStatus === 'pending') {
        bondedCount++;
      }
      
      // Warehouse
      if (location && location !== 'N/A') {
        warehouseCounts[location] = (warehouseCounts[location] || 0) + 1;
        warehouseCount++;
      }
      
      // Buffer
      if (status === 'arrived' && !finalDest) {
        bufferCount++;
      } else if (status === 'aguardando desova' || status === 'pendente') {
          // Backup buffer logic based on previous translations in app
          bufferCount++;
      }

      // Aging mock calculation
      const ageDaysStr = row['AgeDays'] || row['AGE_DAYS'];
      let ageDays = ageDaysStr ? parseInt(String(ageDaysStr), 10) : 0; 
      if (isNaN(ageDays)) ageDays = 0;
      
      if (ageDays <= 7) aging['1-7d']++;
      else if (ageDays <= 15) aging['8-15d']++;
      else if (ageDays <= 30) aging['16-30d']++;
      else aging['30d+']++;

      const waitTimeStr = row['WaitTime'] || row['WAIT_TIME'];
      let waitTime = waitTimeStr ? parseFloat(String(waitTimeStr)) : 0;
      if (isNaN(waitTime)) waitTime = 0;
      totalWaitTime += waitTime;
    });

    const bondedSec = sections.find(s => s.id === 'bonded');
    const warehouseSec = sections.find(s => s.id === 'warehouse');
    const bufferSec = sections.find(s => s.id === 'buffer');
    
    // Use manual if locations are present in that specific section, else fallback to excel
    const finalBonded = (bondedSec && bondedSec.locations.length > 0) ? manualBonded : bondedCount;
    const finalWarehouse = (warehouseSec && warehouseSec.locations.length > 0) ? manualWarehouse : warehouseCount;
    const finalBuffer = (bufferSec && bufferSec.locations.length > 0) ? manualBuffer : bufferCount;
    
    const totalOccupied = finalBonded + finalWarehouse + finalBuffer;
    const finalCapacity = manualCapacity > 0 ? manualCapacity : MAX_CAPACITY;
    const totalCapacity = finalCapacity;
    const utilization = totalCapacity > 0 ? Math.min(100, Math.round((totalOccupied / totalCapacity) * 100)) : 0;

    const distributionData = [
      { name: 'Bonded Area (保税区)', value: finalBonded || 0 },
      { name: 'Warehouse (仓库)', value: finalWarehouse || 0 },
      { name: 'Buffer (缓冲区)', value: finalBuffer || 0 }
    ];

    const finalAgingData = [
      { name: '1-7d', count: manualAging['1-7d'] !== "" ? parseInt(manualAging['1-7d']) || 0 : aging['1-7d'] || 0 },
      { name: '8-15d', count: manualAging['8-15d'] !== "" ? parseInt(manualAging['8-15d']) || 0 : aging['8-15d'] || 0 },
      { name: '16-30d', count: manualAging['16-30d'] !== "" ? parseInt(manualAging['16-30d']) || 0 : aging['16-30d'] || 0 },
      { name: '30d+', count: manualAging['30d+'] !== "" ? parseInt(manualAging['30d+']) || 0 : aging['30d+'] || 0 }
    ];

    let avgStay = finalBuffer > 0 ? (totalWaitTime / bufferCount).toFixed(1) : "0.0";
    if (parseFloat(manualBufferAvgStay) > 0) {
      avgStay = manualBufferAvgStay;
    }

    return {
      totalOccupied,
      totalCapacity,
      utilization,
      distributionData,
      agingData: finalAgingData,
      bufferCount: finalBuffer || 0,
      avgStayTime: avgStay,
    };
  }, [data, sections, manualBufferAvgStay, manualAging]);

  let globalEmpty = 0;
  let globalFull = 0;
  sections.forEach(sec => sec.locations.forEach(loc => {
    globalEmpty += loc.empty;
    globalFull += loc.full;
  }));

  return (
    <div className="flex h-[calc(100vh-64px)] bg-[#f3f4f6]">
      {/* LEFT SIDEBAR - MANUAL CONTROL */}
      <div className="w-[380px] bg-white border-r border-[#e2e8f0] flex flex-col shrink-0 shadow-[4px_0_24px_rgba(0,0,0,0.02)] relative z-10 overflow-hidden">
        <div className="p-6 pb-4 border-b border-[#e2e8f0] flex items-center gap-3">
          <div className="w-10 h-10 bg-[#0f172a] rounded-xl flex items-center justify-center">
            <Box className="text-white w-6 h-6" />
          </div>
          <div>
            <h2 className="text-lg font-bold text-[#0f172a] leading-tight">Storage Inventory</h2>
            <p className="text-xs font-bold text-slate-400 tracking-wider">MANUAL CONTROL</p>
          </div>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-6">
          {sections.map(section => (
            <div key={section.id} className="bg-slate-50 border border-slate-200 rounded-2xl p-4 shadow-sm">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-8 h-8 rounded-full bg-white shadow-sm flex items-center justify-center">
                  {section.icon}
                </div>
                <h3 className="text-sm font-bold text-slate-700 tracking-wider uppercase">{section.title}</h3>
              </div>

              <div className="space-y-4">
                {section.locations.map(loc => (
                  <div key={loc.id} className="bg-white rounded-xl p-3 border border-slate-100 shadow-sm">
                    <h4 className="text-sm font-bold text-[#0f172a] mb-3">{loc.name}</h4>
                    
                    <div className="flex items-center justify-between mb-2 pb-2 border-b border-slate-50">
                       <span className="text-xs font-bold text-slate-400 w-16">CAPACITY</span>
                       <button onClick={() => handleUpdate(section.id, loc.id, 'capacity', -10)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Minus size={14} />
                       </button>
                       <input 
                         type="text" 
                         value={loc.capacity || 0} 
                         onChange={(e) => handeInputMatch(section.id, loc.id, 'capacity', e.target.value)}
                         className="w-16 text-center font-bold text-slate-700 bg-transparent outline-none"
                       />
                       <button onClick={() => handleUpdate(section.id, loc.id, 'capacity', 10)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Plus size={14} />
                       </button>
                    </div>

                    <div className="flex items-center justify-between mb-2">
                       <span className="text-xs font-bold text-slate-400 w-16">EMPTY</span>
                       <button onClick={() => handleUpdate(section.id, loc.id, 'empty', -1)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Minus size={14} />
                       </button>
                       <input 
                         type="text" 
                         value={loc.empty} 
                         onChange={(e) => handeInputMatch(section.id, loc.id, 'empty', e.target.value)}
                         className="w-16 text-center font-bold text-slate-700 bg-transparent outline-none"
                       />
                       <button onClick={() => handleUpdate(section.id, loc.id, 'empty', 1)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Plus size={14} />
                       </button>
                    </div>

                    <div className="flex items-center justify-between">
                       <span className="text-xs font-bold text-[#8b5cf6] w-16">FULL</span>
                       <button onClick={() => handleUpdate(section.id, loc.id, 'full', -1)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Minus size={14} />
                       </button>
                       <input 
                         type="text" 
                         value={loc.full} 
                         onChange={(e) => handeInputMatch(section.id, loc.id, 'full', e.target.value)}
                         className="w-16 text-center font-bold text-[#0f172a] bg-transparent outline-none text-lg"
                       />
                       <button onClick={() => handleUpdate(section.id, loc.id, 'full', 1)} className="w-8 h-8 rounded-full border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-50">
                         <Plus size={14} />
                       </button>
                    </div>
                  </div>
                ))}
                
                {addingSection === section.id ? (
                  <div className="bg-white rounded-xl p-3 border border-slate-200 shadow-sm flex flex-col gap-2">
                    <input
                      autoFocus
                      type="text"
                      placeholder="Location Name"
                      value={newEntryName}
                      onChange={e => setNewEntryName(e.target.value)}
                      onKeyDown={e => {
                        if (e.key === 'Enter') confirmAddLocation(section.id);
                        if (e.key === 'Escape') cancelAddLocation();
                      }}
                      className="w-full text-sm font-bold text-[#0f172a] bg-slate-50 border border-slate-200 rounded p-2 outline-none focus:border-blue-500"
                    />
                    <div className="flex items-center gap-2">
                      <button onClick={() => confirmAddLocation(section.id)} className="flex-1 py-2 bg-blue-600 text-white text-xs font-bold rounded hover:bg-blue-700 transition">ADD</button>
                      <button onClick={cancelAddLocation} className="flex-1 py-2 bg-slate-200 text-slate-600 text-xs font-bold rounded hover:bg-slate-300 transition">CANCEL</button>
                    </div>
                  </div>
                ) : (
                  <button 
                    onClick={() => { setAddingSection(section.id); setNewEntryName(""); }}
                    className="w-full py-3 border border-dashed border-slate-300 rounded-xl text-xs font-bold text-slate-500 flex items-center justify-center gap-2 hover:bg-slate-100 hover:text-slate-700 transition"
                  >
                    <Plus size={14} strokeWidth={3} /> ADD ENTRY
                  </button>
                )}
              </div>
            </div>
          ))}

          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 shadow-sm">
            <h3 className="text-sm font-bold text-slate-700 tracking-wider uppercase mb-4">BUFFER AVG. STAY (DAYS)</h3>
            <div className="bg-white rounded-xl p-3 border border-slate-100 shadow-sm flex items-center justify-between mb-4">
              <span className="text-xs font-bold text-slate-400">DAYS</span>
              <input 
                type="text" 
                value={manualBufferAvgStay} 
                onChange={(e) => setManualBufferAvgStay(e.target.value)}
                className="w-20 text-center font-bold text-[#0f172a] bg-transparent outline-none text-lg border-b border-transparent focus:border-slate-300"
              />
            </div>
            
            <h3 className="text-sm font-bold text-slate-700 tracking-wider uppercase mb-4">INVENTORY AGING</h3>
            <div className="space-y-3">
               {[
                 { key: '1-7d', label: '1-7 Days' },
                 { key: '8-15d', label: '8-15 Days' },
                 { key: '16-30d', label: '16-30 Days' },
                 { key: '30d+', label: '30+ Days' },
               ].map(age => (
                 <div key={age.key} className="bg-white rounded-xl p-3 border border-slate-100 shadow-sm flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-400">{age.label}</span>
                    <input 
                      type="number" 
                      value={manualAging[age.key as keyof typeof manualAging]} 
                      onChange={(e) => setManualAging({...manualAging, [age.key]: e.target.value})}
                      className="w-20 text-center font-bold text-[#0f172a] bg-transparent outline-none text-lg border-b border-transparent focus:border-slate-300"
                      placeholder="Auto"
                    />
                 </div>
               ))}
            </div>
          </div>
        </div>

        <div className="bg-[#0f172a] text-white p-6 shrink-0 shadow-[0_-10px_20px_rgba(0,0,0,0.1)]">
          <div className="flex justify-between mb-4">
            <div>
               <p className="text-[10px] font-bold text-slate-400 tracking-wider mb-1">TOTAL EMPTY</p>
               <p className="text-2xl font-bold text-white">{globalEmpty}</p>
            </div>
            <div className="text-right">
               <p className="text-[10px] font-bold text-[#8b5cf6] tracking-wider mb-1">TOTAL FULL</p>
               <p className="text-2xl font-bold text-white">{globalFull}</p>
            </div>
          </div>
          <div className="pt-4 border-t border-slate-700 flex justify-between items-end">
             <p className="text-[11px] font-bold text-slate-400 tracking-wider">GRAND TOTAL</p>
             <p className="text-3xl font-bold text-[#10b981]">{globalEmpty + globalFull}</p>
          </div>
        </div>
      </div>


      {/* RIGHT SIDE - DASHBOARD */}
      <div className="flex-1 overflow-y-auto p-6 font-sans">
        <div className="max-w-[1000px] mx-auto space-y-6">
          
          {/* Header */}
          <header className="bg-[#003566] rounded-xl p-6 shadow-lg text-white flex justify-between items-center bg-opacity-95 backdrop-blur-md">
            <div>
              <h1 className="text-2xl font-bold tracking-wide">仓储库存看板</h1>
              <h2 className="text-sm text-blue-200 mt-1 uppercase tracking-widest">Storage Inventory Dashboard</h2>
            </div>
            <div className="flex gap-4">
              <div className="bg-[#001d3d] px-4 py-2 rounded-lg border border-blue-800/50 flex flex-col items-end">
                <span className="text-xs text-blue-300">总容量 | Total Cap</span>
                <span className="font-bold text-lg">{metrics.totalCapacity}</span>
              </div>
            </div>
          </header>

          {/* Top KPIs */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {/* Utilization */}
            <div className="bg-white/90 backdrop-blur-sm rounded-xl p-6 shadow-sm border border-slate-200 flex items-center justify-between">
              <div className="flex flex-col">
                <span className="text-sm font-semibold text-slate-800">堆场利用率</span>
                <span className="text-xs text-slate-500 mb-2">Yard Utilization (%)</span>
                <span className="text-3xl font-extrabold text-[#003566]">
                  {metrics.utilization}%
                </span>
              </div>
              <div className="h-16 w-16 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center">
                <Box size={32} />
              </div>
            </div>

            {/* Buffer Total */}
            <div className="bg-white/90 backdrop-blur-sm rounded-xl p-6 shadow-sm border border-slate-200 flex items-center justify-between">
              <div className="flex flex-col">
                <span className="text-sm font-semibold text-slate-800">缓冲区总数</span>
                <span className="text-xs text-slate-500 mb-2">Buffer Total Containers</span>
                <span className="text-3xl font-extrabold text-[#003566]">
                  {metrics.bufferCount}
                </span>
              </div>
              <div className="h-16 w-16 bg-amber-50 text-amber-600 rounded-full flex items-center justify-center">
                <Inbox size={32} />
              </div>
            </div>

            {/* Buffer Avg Wait */}
            <div className="bg-white/90 backdrop-blur-sm rounded-xl p-6 shadow-sm border border-slate-200 flex items-center justify-between">
              <div className="flex flex-col">
                <span className="text-sm font-semibold text-slate-800">平均停留时间</span>
                <span className="text-xs text-slate-500 mb-2">Buffer Avg. Stay (Days)</span>
                <span className="text-3xl font-extrabold text-[#003566]">
                  {metrics.avgStayTime}d
                </span>
              </div>
              <div className="h-16 w-16 bg-emerald-50 text-emerald-600 rounded-full flex items-center justify-center">
                <Package size={32} />
              </div>
            </div>
          </div>

          {/* Charts Row */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Distribution */}
            <div className="bg-white/90 backdrop-blur-sm rounded-xl p-6 shadow-sm border border-slate-200">
               <div className="mb-4">
                <h3 className="text-md font-bold text-slate-800">存储分布</h3>
                <p className="text-xs text-slate-500">Storage Distribution</p>
              </div>
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={metrics.distributionData} layout="vertical" margin={{ top: 5, right: 30, left: 40, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#e2e8f0" />
                    <XAxis type="number" tick={{fontSize: 12, fill: '#64748b'}} />
                    <YAxis dataKey="name" type="category" width={140} tick={{fontSize: 11, fill: '#334155', fontWeight: 500}} />
                    <RechartsTooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                    <Bar dataKey="value" fill="#003566" radius={[0, 4, 4, 0]} barSize={32} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Aging */}
            <div className="bg-white/90 backdrop-blur-sm rounded-xl p-6 shadow-sm border border-slate-200">
              <div className="mb-4">
                <h3 className="text-md font-bold text-slate-800">库存库龄</h3>
                <p className="text-xs text-slate-500">Inventory Aging</p>
              </div>
              <div className="h-72">
                 <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={metrics.agingData} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                    <XAxis dataKey="name" tick={{fontSize: 12, fill: '#64748b'}} axisLine={false} tickLine={false} dy={10} />
                    <YAxis tick={{fontSize: 12, fill: '#64748b'}} axisLine={false} tickLine={false} dx={-10} />
                    <RechartsTooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                    <Bar dataKey="count" fill="#38bdf8" radius={[4, 4, 0, 0]} barSize={48}>
                      {
                        metrics.agingData.map((entry, index) => {
                          const colors = ['#22d3ee', '#22d3ee', '#22d3ee', '#22d3ee'];
                          return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
                        })
                      }
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
};

let root: any = null;

export function mountStorageInventory(container: HTMLElement, data: any[]) {
  if (!root) {
    root = createRoot(container);
  }
  root.render(<StorageInventory data={data} />);
}

