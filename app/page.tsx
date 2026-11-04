// src/app/page.tsx
// ==========================================
// VUE PRINCIPALE : Dashboard, Graphes et Filtres Périodiques Avancés (2026)
// ==========================================
'use client';

import { useState, useMemo, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';

// Données par défaut initiales
const defaultExcelData = [
  { id: '1', date: '02.11.2026', time: '01:48:00', zeit: '02.11.2026 01:48:00', kanalKunde: 'KYZ Martin', buchung: 'Live-Ticket Abgabe', betrag: 349.00 },
  { id: '2', date: '03.11.2026', time: '08:13:00', zeit: '02.11.2026 08:13:00', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 300.00 },
  { id: '3', date: '02.11.2026', time: '08:18:15', zeit: '02.11.2026 08:18:15', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 300.00 },
  { id: '4', date: '02.11.2026', time: '08:26:40', zeit: '02.11.2026 08:26:40', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '5', date: '02.11.2026', time: '08:32:10', zeit: '02.11.2026 08:32:10', kanalKunde: 'KYZ Mokolo', buchung: 'Live-Ticket Abgabe', betrag: 200.00 },
  { id: '6', date: '02.11.2026', time: '08:33:55', zeit: '02.11.2026 08:33:55', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '7', date: '04.11.2026', time: '08:35:20', zeit: '02.11.2026 08:35:20', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '8', date: '02.11.2026', time: '08:41:05', zeit: '02.11.2026 08:41:05', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '9', date: '02.11.2026', time: '08:47:30', zeit: '02.11.2026 08:47:30', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '10', date: '02.11.2026', time: '08:47:45', zeit: '02.11.2026 08:47:45', kanalKunde: 'KYZ Ndokotti', buchung: 'Ticketabgabe', betrag: 800.00 },
  { id: '11', date: '02.11.2026', time: '08:57:12', zeit: '02.11.2026 08:57:12', kanalKunde: 'KYZ Ndokotti', buchung: 'Ticketabgabe', betrag: 300.00 },
  { id: '12', date: '02.11.2026', time: '08:57:50', zeit: '02.11.2026 08:57:50', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 270.00 },
  { id: '13', date: '02.11.2026', time: '09:05:00', zeit: '02.11.2026 09:05:00', kanalKunde: 'KYZ Mokolo', buchung: 'Ticketabgabe', betrag: 200.00 },
  { id: '14', date: '02.11.2026', time: '09:13:22', zeit: '02.11.2026 09:13:22', kanalKunde: 'KYZ Bonaberi', buchung: 'Live-Ticket Abgabe', betrag: 200.00 },
  { id: '15', date: '02.11.2026', time: '09:27:10', zeit: '02.11.2026 09:27:10', kanalKunde: 'KYZ Bonaberi', buchung: 'Live-Ticket Abgabe', betrag: 200.00 },
  { id: '16', date: '02.11.2026', time: '09:29:40', zeit: '02.11.2026 09:29:40', kanalKunde: 'KYZ Logbessou', buchung: 'Ticketabgabe', betrag: 200.00 },
];

const hexColors = [
  '#4F46E5', '#7C3AED', '#059669', '#D97706', '#E11D48',
  '#0284C7', '#C026D3', '#0D9488', '#EA580C', '#0891B2',
  '#3B82F6', '#A855F7', '#EC4899', '#84CC16', '#EAB308', '#EF4444'
];

export default function HomePage() {
  const [activeMenu, setActiveMenu] = useState<string>('dashboard');
  const [activeTab, setActiveTab] = useState<string>('Tous');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [tableData, setTableData] = useState(defaultExcelData);

  // États pour la Période Initiale
  const [startYear, setStartYear] = useState<string>('Tous');
  const [startMonth, setStartMonth] = useState<string>('Tous');
  const [startDay, setStartDay] = useState<string>('Tous');
  const [startTimeInput, setStartTimeInput] = useState<string>('');

  // États pour la Période de Fin
  const [endYear, setEndYear] = useState<string>('Tous');
  const [endMonth, setEndMonth] = useState<string>('Tous');
  const [endDay, setEndDay] = useState<string>('Tous');
  const [endTimeInput, setEndTimeInput] = useState<string>('');

  // Génération de la liste des années (De 2026 en descendant)
  const yearsList = useMemo(() => {
    const currentYear = 2026;
    const years = [];
    for (let y = currentYear; y >= 2020; y--) {
      years.push(String(y));
    }
    return years;
  }, []);

  // Liste des 12 mois
  const monthsList = [
    { num: '01', name: '01 - Janvier' },
    { num: '02', name: '02 - Février' },
    { num: '03', name: '03 - Mars' },
    { num: '04', name: '04 - Avril' },
    { num: '05', name: '05 - Mai' },
    { num: '06', name: '06 - Juin' },
    { num: '07', name: '07 - Juillet' },
    { num: '08', name: '08 - Août' },
    { num: '09', name: '09 - Septembre' },
    { num: '10', name: '10 - Octobre' },
    { num: '11', name: '11 - Novembre' },
    { num: '12', name: '12 - Décembre' },
  ];

  // Calcul du nombre de jours dynamique selon l'année et le mois choisis
  const getDaysInMonthList = (yearStr: string, monthStr: string) => {
    if (monthStr === 'Tous') return Array.from({ length: 31 }, (_, i) => String(i + 1).padStart(2, '0'));
    const y = yearStr === 'Tous' ? 2026 : parseInt(yearStr, 10);
    const m = parseInt(monthStr, 10);
    const daysCount = new Date(y, m, 0).getDate();
    return Array.from({ length: daysCount }, (_, i) => String(i + 1).padStart(2, '0'));
  };

  const startDaysList = useMemo(() => getDaysInMonthList(startYear, startMonth), [startYear, startMonth]);
  const endDaysList = useMemo(() => getDaysInMonthList(endYear, endMonth), [endYear, endMonth]);

  // Parsing automatique de la date et de l'heure HH:MM:SS
  const parseDateTimeString = (rawStr: string) => {
    let datePart = '02.11.2026';
    let timePart = '01:48:00';

    if (rawStr.includes(' ')) {
      const parts = rawStr.split(' ');
      datePart = parts[0];
      timePart = parts[1];
    } else if (rawStr.includes(':')) {
      timePart = rawStr;
    } else if (rawStr.includes('.')) {
      datePart = rawStr;
    }

    const timeSegments = timePart.split(':');
    if (timeSegments.length === 2) {
      timePart = `${timePart}:00`;
    }

    return {
      date: datePart,
      time: timePart,
      zeit: `${datePart} ${timePart}`
    };
  };

  // Chargement de fichier Excel / CSV / HTML
  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    if (file.name.endsWith('.csv') || file.name.endsWith('.html') || file.name.endsWith('.htm')) {
      reader.onload = (event) => parseRawTextContent(event.target?.result as string);
      reader.readAsText(file);
    } else {
      reader.onload = (event) => {
        try {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const uint8Array = new Uint8Array(arrayBuffer);
          const textDecoder = new TextDecoder('utf-8');
          const fileHeader = textDecoder.decode(uint8Array.slice(0, 100));

          if (fileHeader.includes('<html') || fileHeader.includes('<table')) {
            parseRawTextContent(textDecoder.decode(uint8Array));
            return;
          }

          const workbook = XLSX.read(uint8Array, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 1 });
          processJsonRows(jsonData);
        } catch (err) {
          reader.onload = (ev) => parseRawTextContent(ev.target?.result as string);
          reader.readAsText(file);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const parseRawTextContent = (text: string) => {
    try {
      if (text.includes('<tr') || text.includes('<table')) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(text, 'text/html');
        const rows = doc.querySelectorAll('tr');
        const matrix: string[][] = [];
        rows.forEach((tr) => {
          const cells = tr.querySelectorAll('td, th');
          const rowData: string[] = [];
          cells.forEach((cell) => rowData.push(cell.textContent?.trim() || ''));
          if (rowData.some(cell => cell !== '')) matrix.push(rowData);
        });
        if (matrix.length > 0) processJsonRows(matrix);
        return;
      }

      const lines = text.split(/\r\n|\n/);
      const matrix = lines.map(line => line.split(/[,;\t]/).map(c => c.replace(/"/g, '').trim()));
      processJsonRows(matrix);
    } catch (e) {
      alert("Impossible de lire ce format de fichier.");
    }
  };

  const processJsonRows = (jsonData: any[][]) => {
    let headerRowIndex = 0;
    for (let i = 0; i < jsonData.length; i++) {
      const rowStr = JSON.stringify(jsonData[i]).toLowerCase();
      if (rowStr.includes('zeit') || rowStr.includes('kanal') || rowStr.includes('buchung') || rowStr.includes('betrag') || rowStr.includes('date')) {
        headerRowIndex = i;
        break;
      }
    }

    const headers = jsonData[headerRowIndex] || [];
    const rows = jsonData.slice(headerRowIndex + 1);
    let parsedRows: any[] = [];

    rows.forEach((row, index) => {
      if (!row || row.length === 0) return;

      let rawZeit = '';
      let kanalKunde = '';
      let buchung = '';
      let betrag = 0;

      headers.forEach((header: any, colIdx: number) => {
        const h = String(header).toLowerCase().trim();
        const val = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';

        if (h.includes('zeit') || h.includes('date') || h.includes('heure')) {
          rawZeit = val;
        } else if (h.includes('kanal') || h.includes('kunde') || h.includes('client')) {
          kanalKunde = val;
        } else if (h.includes('buchung') || h.includes('operation') || h.includes('action')) {
          buchung = val;
        } else if (h.includes('betrag') || h.includes('montant') || h.includes('prix')) {
          betrag = parseFloat(val.replace(/[^0-9.-]+/g, '')) || 0;
        }
      });

      if (!rawZeit && row[0]) rawZeit = String(row[0]).trim();
      if (!kanalKunde && row[1]) kanalKunde = String(row[1]).trim();
      if (!buchung && row[2]) buchung = String(row[2]).trim();
      if (betrag === 0 && row[3]) {
        betrag = parseFloat(String(row[3]).replace(/[^0-9.-]+/g, '')) || 0;
      }

      const dt = parseDateTimeString(rawZeit || '02.11.2026 01:48:00');

      if (kanalKunde || buchung || betrag > 0) {
        parsedRows.push({
          id: String(index + 1),
          date: dt.date,
          time: dt.time,
          zeit: dt.zeit,
          kanalKunde: kanalKunde || 'Client Inconnu',
          buchung: buchung || 'Ticketabgabe',
          betrag: betrag,
        });
      }
    });

    if (parsedRows.length > 0) setTableData(parsedRows);
  };

  // Filtrage combiné (Recherche d'éléments + Intervalle périodique)
  const filteredData = useMemo(() => {
    return tableData.filter((item) => {
      const matchesSearch = Object.values(item).some((val) =>
        String(val).toLowerCase().includes(searchTerm.toLowerCase())
      );

      const [dDay, dMonth, dYear] = item.date.split('.');
      const dTime = item.time;

      // Filtrage par Période Initiale
      if (startYear !== 'Tous' && dYear < startYear) return false;
      if (startMonth !== 'Tous' && dMonth < startMonth) return false;
      if (startDay !== 'Tous' && dDay < startDay) return false;
      if (startTimeInput.trim() !== '' && dTime < startTimeInput.trim()) return false;

      // Filtrage par Période de Fin
      if (endYear !== 'Tous' && dYear > endYear) return false;
      if (endMonth !== 'Tous' && dMonth > endMonth) return false;
      if (endDay !== 'Tous' && dDay > endDay) return false;
      if (endTimeInput.trim() !== '' && dTime > endTimeInput.trim()) return false;

      return matchesSearch;
    });
  }, [searchTerm, tableData, startYear, startMonth, startDay, startTimeInput, endYear, endMonth, endDay, endTimeInput]);

  const totalOperations = filteredData.length;
  const amounts = filteredData.map((d) => d.betrag);
  const totalBetrag = amounts.reduce((acc, val) => acc + val, 0);

  const chartData = useMemo(() => {
    if (activeTab === 'Zeit') {
      return filteredData.map((item, idx) => ({
        label: item.time,
        value: item.betrag,
        sub: item.kanalKunde,
        color: hexColors[idx % hexColors.length],
      }));
    } else if (activeTab === 'Kanal/Kunde') {
      const grouped: { [key: string]: number } = {};
      filteredData.forEach((item) => {
        grouped[item.kanalKunde] = (grouped[item.kanalKunde] || 0) + item.betrag;
      });
      return Object.entries(grouped).map(([key, val], idx) => ({
        label: key,
        value: val,
        sub: 'Client',
        color: hexColors[idx % hexColors.length],
      }));
    } else if (activeTab === 'Buchung') {
      const grouped: { [key: string]: number } = {};
      filteredData.forEach((item) => {
        grouped[item.buchung] = (grouped[item.buchung] || 0) + item.betrag;
      });
      return Object.entries(grouped).map(([key, val], idx) => ({
        label: key,
        value: val,
        sub: 'Opération',
        color: hexColors[idx % hexColors.length],
      }));
    } else {
      return filteredData.map((item, idx) => ({
        label: `${item.kanalKunde}`,
        value: item.betrag,
        sub: item.buchung,
        color: hexColors[idx % hexColors.length],
      }));
    }
  }, [filteredData, activeTab]);

  const pieSlices = useMemo(() => {
    const sum = chartData.reduce((acc, d) => acc + d.value, 0) || 1;
    let cumulativeAngle = 0;

    return chartData.map((item) => {
      const percentage = (item.value / sum) * 100;
      const angle = (item.value / sum) * 360;
      const startAngle = cumulativeAngle;
      const endAngle = cumulativeAngle + angle;
      cumulativeAngle += angle;

      const x1 = 150 + 100 * Math.cos((Math.PI * (startAngle - 90)) / 180);
      const y1 = 150 + 100 * Math.sin((Math.PI * (startAngle - 90)) / 180);
      const x2 = 150 + 100 * Math.cos((Math.PI * (endAngle - 90)) / 180);
      const y2 = 150 + 100 * Math.sin((Math.PI * (endAngle - 90)) / 180);

      const largeArcFlag = angle > 180 ? 1 : 0;
      const pathData = `M 150 150 L ${x1} ${y1} A 100 100 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;

      const midAngle = startAngle + angle / 2;
      const labelX = 150 + 125 * Math.cos((Math.PI * (midAngle - 90)) / 180);
      const labelY = 150 + 125 * Math.sin((Math.PI * (midAngle - 90)) / 180);

      return {
        ...item,
        percentage: percentage.toFixed(2),
        pathData,
        labelX,
        labelY,
      };
    });
  }, [chartData]);

  return (
    <div className="min-h-screen bg-zinc-50 dark:bg-black text-black dark:text-zinc-50 flex">
      
      {/* SIDEBAR */}
      <aside className="w-64 border-r border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-950 p-6 hidden md:flex flex-col justify-between">
        <div>
          <div className="flex items-center gap-3 mb-10">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center text-white font-bold">E</div>
            <span className="font-bold text-lg tracking-tight">Essai_1</span>
          </div>

          <nav className="space-y-1">
            <button onClick={() => setActiveMenu('dashboard')} className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl font-medium text-sm transition-colors ${activeMenu === 'dashboard' ? 'bg-indigo-600 text-white' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-900'}`}>Tableau de bord</button>
            <button onClick={() => setActiveMenu('diagrammes')} className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl font-medium text-sm transition-colors ${activeMenu === 'diagrammes' ? 'bg-indigo-600 text-white' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-900'}`}>Graphes en diagrammes</button>
            <button onClick={() => setActiveMenu('bandes')} className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl font-medium text-sm transition-colors ${activeMenu === 'bandes' ? 'bg-indigo-600 text-white' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-900'}`}>Graphes en bandes</button>
            <button onClick={() => setActiveMenu('secteurs')} className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl font-medium text-sm transition-colors ${activeMenu === 'secteurs' ? 'bg-indigo-600 text-white' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-900'}`}>Graphes en secteurs</button>
            <button onClick={() => setActiveMenu('parametres')} className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl font-medium text-sm transition-colors ${activeMenu === 'parametres' ? 'bg-indigo-600 text-white' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-900'}`}>Paramètres</button>
          </nav>
        </div>

        <div className="pt-4 border-t border-zinc-200 dark:border-zinc-800 text-xs text-zinc-500">
          Connecté (Mode Gratuit)
        </div>
      </aside>

      {/* CONTENU PRINCIPAL */}
      <main className="flex-1 flex flex-col min-w-0">
        
        <header className="h-20 border-b border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-950 px-8 flex items-center justify-between">
          <h1 className="text-xl font-bold tracking-tight uppercase">
            {activeMenu === 'dashboard' && 'TABLEAU DE BORD (DONNÉES ACTIVES)'}
            {activeMenu === 'diagrammes' && 'Graphes en diagrammes (Histogramme unifié)'}
            {activeMenu === 'bandes' && 'Graphes en bandes'}
            {activeMenu === 'secteurs' && 'Graphes en secteurs (Pie Chart unifié)'}
            {activeMenu === 'parametres' && 'Paramètres du système'}
          </h1>
          <div className="flex items-center gap-4">
            <div className="w-9 h-9 rounded-full bg-indigo-100 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400 font-bold flex items-center justify-center text-sm">
              TIAM
            </div>
          </div>
        </header>

        <div className="p-8 max-w-7xl w-full mx-auto space-y-8">
          
          {/* DIV PRINCIPALE DE RECHERCHE */}
          <div className="flex flex-col gap-6 bg-white dark:bg-zinc-950 p-6 rounded-2xl border border-zinc-200 dark:border-zinc-800 shadow-sm">
            
            {/* Ligne Supérieure : Recherche globale & Importation */}
            <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 pb-4 border-b border-zinc-100 dark:border-zinc-900">
              <div className="w-full lg:w-auto">
                <h2 className="text-sm font-semibold text-zinc-700 dark:text-zinc-300">Recherche d'éléments</h2>
                <input
                  type="text"
                  placeholder="Rechercher (ex: Mokolo)..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="mt-1 w-full sm:w-80 px-4 py-2 text-sm rounded-xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-black focus:outline-none focus:border-indigo-600 font-medium"
                />
              </div>

              <div className="flex items-center gap-3 bg-zinc-50 dark:bg-black p-3 rounded-xl border border-zinc-200 dark:border-zinc-800 w-full lg:w-auto">
                <div>
                  <p className="text-xs font-semibold text-zinc-700 dark:text-zinc-300">Charger tout fichier Excel / CSV / HTML</p>
                  <p className="text-[10px] text-zinc-500">.xlsx, .xls, .csv pris en charge</p>
                </div>
                <input 
                  type="file" 
                  accept=".csv, .xls, .xlsx, .html, .htm"
                  onChange={handleFileUpload}
                  className="text-xs text-zinc-500 file:mr-4 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-indigo-600 file:text-white hover:file:bg-indigo-700 cursor-pointer"
                />
              </div>
            </div>

            {/* SECTION : RECHERCHE PÉRIODIQUE */}
            <div>
              <h3 className="text-sm font-bold text-indigo-600 dark:text-indigo-400 mb-3">Recherche périodique</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* Div Période Initiale */}
                <div className="p-4 rounded-xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900/40 space-y-3">
                  <p className="text-xs font-semibold text-zinc-700 dark:text-zinc-300">Période initiale</p>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste d'année</label>
                      <select value={startYear} onChange={(e) => setStartYear(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {yearsList.map(y => <option key={y} value={y}>{y}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste des mois</label>
                      <select value={startMonth} onChange={(e) => setStartMonth(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {monthsList.map(m => <option key={m.num} value={m.num}>{m.name}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste des jours</label>
                      <select value={startDay} onChange={(e) => setStartDay(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {startDaysList.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Horaire (HH:MM:SS)</label>
                      <input 
                        type="text"
                        placeholder="00:00:00"
                        value={startTimeInput}
                        onChange={(e) => setStartTimeInput(e.target.value)}
                        className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-mono focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                  </div>
                </div>

                {/* Div Période de Fin */}
                <div className="p-4 rounded-xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900/40 space-y-3">
                  <p className="text-xs font-semibold text-zinc-700 dark:text-zinc-300">Période de fin</p>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste d'année</label>
                      <select value={endYear} onChange={(e) => setEndYear(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {yearsList.map(y => <option key={y} value={y}>{y}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste des mois</label>
                      <select value={endMonth} onChange={(e) => setEndMonth(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {monthsList.map(m => <option key={m.num} value={m.num}>{m.name}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Liste des jours</label>
                      <select value={endDay} onChange={(e) => setEndDay(e.target.value)} className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-medium">
                        <option value="Tous">Tous</option>
                        {endDaysList.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] text-zinc-500">Horaire (HH:MM:SS)</label>
                      <input 
                        type="text"
                        placeholder="23:59:59"
                        value={endTimeInput}
                        onChange={(e) => setEndTimeInput(e.target.value)}
                        className="w-full mt-1 p-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black font-mono focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                  </div>
                </div>

              </div>
            </div>

          </div>

          {/* CARDS D'INDICATEURS */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div onClick={() => setActiveTab('Zeit')} className={`cursor-pointer p-5 rounded-2xl border transition-all ${activeTab === 'Zeit' ? 'border-indigo-600 bg-indigo-50/50 dark:bg-indigo-950/30' : 'bg-white dark:bg-black border-zinc-200 dark:border-zinc-800'}`}>
              <p className="text-xs font-medium text-zinc-500">Zeit (Plage)</p>
              <p className="text-sm font-bold mt-2">Dynamique</p>
            </div>
            <div onClick={() => setActiveTab('Kanal/Kunde')} className={`cursor-pointer p-5 rounded-2xl border transition-all ${activeTab === 'Kanal/Kunde' ? 'border-indigo-600 bg-indigo-50/50 dark:bg-indigo-950/30' : 'bg-white dark:bg-black border-zinc-200 dark:border-zinc-800'}`}>
              <p className="text-xs font-medium text-zinc-500">Kanal/Kunde</p>
              <p className="text-lg font-bold mt-2">{totalOperations} Clients</p>
            </div>
            <div onClick={() => setActiveTab('Buchung')} className={`cursor-pointer p-5 rounded-2xl border transition-all ${activeTab === 'Buchung' ? 'border-indigo-600 bg-indigo-50/50 dark:bg-indigo-950/30' : 'bg-white dark:bg-black border-zinc-200 dark:border-zinc-800'}`}>
              <p className="text-xs font-medium text-zinc-500">Buchung</p>
              <p className="text-lg font-bold mt-2">{totalOperations} Opérations</p>
            </div>
            <div onClick={() => setActiveTab('Betrag')} className={`cursor-pointer p-5 rounded-2xl border transition-all ${activeTab === 'Betrag' ? 'border-indigo-600 bg-indigo-50/50 dark:bg-indigo-950/30' : 'bg-white dark:bg-black border-zinc-200 dark:border-zinc-800'}`}>
              <p className="text-xs font-medium text-zinc-500">Betrag (Total XAF)</p>
              <p className="text-base font-bold text-indigo-600 dark:text-indigo-400 mt-2">{totalBetrag.toLocaleString()} XAF</p>
            </div>
            <div onClick={() => setActiveTab('Tous')} className={`cursor-pointer p-5 rounded-2xl border transition-all ${activeTab === 'Tous' ? 'border-indigo-600 bg-indigo-600 text-white' : 'bg-white dark:bg-black border-zinc-200 dark:border-zinc-800'}`}>
              <p className={`text-xs font-medium ${activeTab === 'Tous' ? 'text-indigo-100' : 'text-zinc-500'}`}>Vue Globale</p>
              <p className="text-lg font-bold mt-2">Tous</p>
            </div>
          </div>

          {/* TABLEAU DE BORD */}
          {activeMenu === 'dashboard' && (
            <section className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-lg font-semibold">
                  {activeTab === 'Tous' ? 'Tableau de Bord - Données Actives' : `Colonne active : ${activeTab}`}
                </h2>
                <span className="text-xs text-zinc-500">{totalOperations} lignes filtrées</span>
              </div>

              <div className="max-h-[400px] overflow-auto border border-zinc-200 dark:border-zinc-800 rounded-xl">
                <table className="w-full text-left border-collapse min-w-[700px]">
                  <thead className="sticky top-0 z-10">
                    <tr className="bg-indigo-600 text-white text-xs uppercase tracking-wider">
                      {(activeTab === 'Tous' || activeTab === 'Zeit') && <th className="p-4 font-bold">Zeit (Date & Heure)</th>}
                      {(activeTab === 'Tous' || activeTab === 'Kanal/Kunde') && <th className="p-4 font-bold">Kanal/Kunde</th>}
                      {(activeTab === 'Tous' || activeTab === 'Buchung') && <th className="p-4 font-bold">Buchung</th>}
                      {(activeTab === 'Tous' || activeTab === 'Betrag') && <th className="p-4 font-bold">Betrag</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-zinc-100 dark:divide-zinc-900 text-sm">
                    {filteredData.map((row) => (
                      <tr key={row.id} className="hover:bg-zinc-50 dark:hover:bg-zinc-900 transition-colors">
                        {(activeTab === 'Tous' || activeTab === 'Zeit') && (<td className="p-4 font-mono text-xs">{row.zeit}</td>)}
                        {(activeTab === 'Tous' || activeTab === 'Kanal/Kunde') && (<td className="p-4">{row.kanalKunde}</td>)}
                        {(activeTab === 'Tous' || activeTab === 'Buchung') && (<td className="p-4 text-indigo-600 dark:text-indigo-400 font-medium">{row.buchung}</td>)}
                        {(activeTab === 'Tous' || activeTab === 'Betrag') && (<td className="p-4 font-semibold">{row.betrag.toFixed(2)} XAF</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* DIAGRAMMES */}
          {activeMenu === 'diagrammes' && (
            <section className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm space-y-6">
              <h2 className="text-lg font-semibold">Histogramme unifié : {activeTab}</h2>
              <div className="p-6 rounded-2xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900/20">
                <div className="h-64 flex items-end gap-3 pt-6 pb-2 px-2 border-b border-l border-zinc-300 dark:border-zinc-700 overflow-x-auto">
                  {chartData.map((item, idx) => {
                    const heightPct = Math.max(15, Math.min(100, (item.value / Math.max(...chartData.map(d => d.value), 1)) * 100));
                    return (
                      <div key={idx} className="flex-1 flex flex-col items-center min-w-[50px] h-full justify-end group relative">
                        <span className="absolute -top-6 text-[10px] font-bold text-indigo-600 dark:text-indigo-400 opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap">
                          {item.value} XAF
                        </span>
                        <div className="w-full rounded-t-lg transition-all shadow-sm hover:brightness-110 flex items-center justify-center" style={{ height: `${heightPct}%`, backgroundColor: item.color }}>
                          <span className="text-[9px] text-white font-bold rotate-90 sm:rotate-0 truncate px-1">{item.value}</span>
                        </div>
                        <span className="text-[10px] font-medium text-zinc-600 dark:text-zinc-400 mt-2 truncate w-full text-center" title={item.label}>{item.label}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </section>
          )}

          {/* SECTEURS */}
          {activeMenu === 'secteurs' && (
            <section className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm space-y-6">
              <div className="flex justify-between items-center">
                <h2 className="text-lg font-semibold">Graphique en secteur (Camembert) : {activeTab}</h2>
                <span className="text-xs text-zinc-500">Répartition par catégorie</span>
              </div>
              <div className="p-8 rounded-2xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900/20 flex flex-col items-center justify-center">
                <div className="relative w-full max-w-[500px] h-[360px] flex items-center justify-center">
                  <svg viewBox="0 0 300 300" className="w-full h-full drop-shadow-md">
                    {pieSlices.map((slice, idx) => (
                      <g key={idx}>
                        <path d={slice.pathData} fill={slice.color} stroke="#ffffff" strokeWidth="1.5" className="hover:opacity-90 transition-opacity cursor-pointer" />
                        <text x={slice.labelX} y={slice.labelY} fill="currentColor" fontSize="9" fontWeight="bold" textAnchor="middle" dominantBaseline="central" className="text-zinc-800 dark:text-zinc-200">
                          {slice.label.length > 12 ? `${slice.label.slice(0, 10)}...` : slice.label}
                        </text>
                        <text x={slice.labelX} y={slice.labelY + 11} fill={slice.color} fontSize="8" fontWeight="bold" textAnchor="middle" dominantBaseline="central">
                          {slice.value} XAF ({slice.percentage}%)
                        </text>
                      </g>
                    ))}
                  </svg>
                </div>
              </div>
            </section>
          )}

          {activeMenu === 'bandes' && (
            <section className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm space-y-6">
              <h2 className="text-lg font-semibold">Graphes en bandes</h2>
              <div className="space-y-4">
                {filteredData.map((item, idx) => {
                  const widthPercent = Math.min(100, Math.max(20, (item.betrag / Math.max(...amounts, 1)) * 100));
                  return (
                    <div key={item.id} className="p-4 rounded-xl border border-zinc-200 dark:border-zinc-800">
                      <div className="flex justify-between text-sm mb-1">
                        <span className="font-medium">{item.kanalKunde} ({item.buchung})</span>
                        <span className="font-bold text-indigo-600">{item.betrag} XAF</span>
                      </div>
                      <div className="w-full bg-zinc-100 dark:bg-zinc-900 h-4 rounded-full overflow-hidden">
                        <div className="h-full rounded-full bg-indigo-600" style={{ width: `${widthPercent}%` }}></div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>
          )}

          {activeMenu === 'parametres' && (
            <section className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm space-y-4">
              <h2 className="text-lg font-semibold">Paramètres de l'application</h2>
              <p className="text-sm text-zinc-500">Configuration générale de l'interface.</p>
            </section>
          )}

        </div>
      </main>
    </div>
  );
}