import React, { useState, useEffect, useMemo } from 'react';
import { 
  ShoppingCart, 
  Glasses, 
  Package, 
  Wrench, 
  User, 
  ClipboardList, 
  AlertTriangle, 
  CheckCircle, 
  Truck, 
  ArrowRight,
  Plus,
  History,
  Store,
  HardHat,
  Banknote,
  CreditCard,
  QrCode,
  Receipt,
  X,
  Lock,
  Settings,
  Users,
  Building,
  PackagePlus,
  ClipboardCheck,
  ArrowDownCircle,
  ArrowUpCircle,
  FileText,
  Send,
  Database,
  FileSpreadsheet,
  Globe,
  Shield,
  Clock,
  Crown,
  Calendar,
  Code,
  Timer,
  Menu // Ikon Menu untuk mobile jika diperlukan nanti
} from 'lucide-react';

// --- DATA DUMMY AWAL (INITIAL DATA) ---

const INITIAL_BRANCHES = [
  { id: 'B1', name: 'Cabang Melawai (Sales Only)', type: 'SALES' },
  { id: 'B2', name: 'Cabang Sudirman (Sales Only)', type: 'SALES' },
  { id: 'B3', name: 'Cabang Kemang (Sales Only)', type: 'SALES' },
  { id: 'B4', name: 'Cabang Pusat (PRODUKSI)', type: 'PRODUCTION' },
  { id: 'B5', name: 'Cabang Barat (PRODUKSI)', type: 'PRODUCTION' }
];

const INITIAL_TECHNICIANS = [
  { id: 'T1', name: 'Andi (Senior Tech)' },
  { id: 'T2', name: 'Budi (Lens Specialist)' },
  { id: 'T3', name: 'Citra (QC & Fitting)' },
  { id: 'T4', name: 'Dewi (Junior Tech)' }
];

// Inventory Pusat (Gudang Master)
const INITIAL_INVENTORY = [
  { id: 'F001', name: 'Frame Rayban Aviator', type: 'frame', stock: 100, price: 1500000 },
  { id: 'F002', name: 'Frame Gucci Acetate', type: 'frame', stock: 50, price: 2500000 },
  { id: 'L001', name: 'Lensa Essilor SV -2.00', type: 'lens', stock: 200, price: 500000 },
  { id: 'L002', name: 'Lensa Hoya BlueControl', type: 'lens', stock: 150, price: 750000 },
  { id: 'P001', name: 'Nose Pad Silica', type: 'part', stock: 1000, price: 10000 },
  { id: 'P002', name: 'Screw Set Optik', type: 'part', stock: 2000, price: 5000 },
];

// Stok Awal di Cabang (Sedikit untuk simulasi request)
const INITIAL_BRANCH_STOCK = {
  'B1': { 'F001': 5, 'F002': 2, 'L001': 10, 'L002': 5 },
  'B2': { 'F001': 3, 'F002': 3, 'L001': 5, 'L002': 5 },
  'B3': { 'F001': 4, 'F002': 1, 'L001': 8, 'L002': 4 },
  'B4': { 'P001': 50, 'P002': 100, 'L001': 20, 'F001': 2 }, // Produksi butuh part
  'B5': { 'P001': 30, 'P002': 50, 'L001': 15, 'F001': 1 }
};

const INITIAL_ORDERS = [
  {
    id: 'ORD-001',
    customer: 'Budi Santoso',
    branchId: 'B1',
    productionBranchId: 'B4',
    status: 'IN_PROGRESS',
    technician: 'Andi (Senior Tech)', 
    items: [{ itemId: 'F001', name: 'Frame Rayban Aviator', price: 1500000 }],
    totalPrice: 1500000,
    paymentMethod: 'CASH',
    prescription: {
      od: { sph: '-2.00', cyl: '-0.50', axis: '90', add: '' },
      os: { sph: '-1.75', cyl: '0.00', axis: '0', add: '' },
      pd: '64'
    },
    logs: [
      { date: '2023-10-25 10:00', msg: 'Order dibuat di Cabang Melawai', type: 'info' },
      { date: '2023-10-25 10:05', msg: 'Masuk antrian produksi Cabang Pusat', type: 'system' },
      { date: '2023-10-25 11:00', msg: 'Mulai dikerjakan oleh Teknisi Andi', type: 'progress' }
    ],
    partsUsed: []
  }
];

// Helper untuk tambah tanggal
const addTime = (baseDateStr, type, amount) => {
  const base = baseDateStr ? new Date(baseDateStr) : new Date();
  if (isNaN(base.getTime())) return new Date().toISOString().slice(0, 16);
  const result = new Date(base);
  if (type === 'days') result.setDate(result.getDate() + amount);
  if (type === 'months') result.setMonth(result.getMonth() + amount);
  if (type === 'years') result.setFullYear(result.getFullYear() + amount);
  const offset = result.getTimezoneOffset() * 60000;
  return (new Date(result - offset)).toISOString().slice(0, 16);
};

// Konfigurasi Awal Sistem (Disesuaikan dengan Referensi)
const INITIAL_APP_CONFIG = {
  dbType: 'LOCAL', 
  gsheetUrl: '',
  // Config Trial & Langganan Baru
  trialEnabled: true,
  trialEndDate: addTime(new Date(), 'days', 14), // Default 14 hari
  subscriptionEnabled: false,
  subscriptionEndDate: '',
  isActive: true
};

const GAS_TEMPLATE = `// --- GOOGLE APPS SCRIPT CODE ---
// Salin kode ini ke Editor Skrip Google Sheet Anda (Extensions > Apps Script)

function doGet(e) {
  // Setup Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventory') || ss.insertSheet('Inventory');
  
  // Ambil Data
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Hapus header
  
  const result = data.map(row => {
    return {
      id: row[0],
      name: row[1],
      type: row[2],
      stock: row[3],
      price: row[4]
    };
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    data: result
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Logika Simpan Data akan di sini (Sesuai kebutuhan)
    // Contoh sederhana:
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data berhasil disinkronisasi'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
`;

// --- UTILS ---
const formatRupiah = (number) => {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR' }).format(number);
};

const getCurrentTime = () => new Date().toLocaleString('id-ID');

// Helper Component: Countdown
const CountdownDisplay = ({ targetDate, label = "Sisa Waktu" }) => {
  const [timeLeft, setTimeLeft] = useState('');
  useEffect(() => {
    const calculateTime = () => {
      const now = new Date();
      const end = new Date(targetDate);
      const diff = end - now;
      if (diff <= 0) return 'Kadaluarsa';
      const days = Math.floor(diff / (1000 * 60 * 60 * 24));
      const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
      const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
      return `${days} Hari ${hours} Jam ${minutes} Menit`;
    };
    setTimeLeft(calculateTime());
    const interval = setInterval(() => setTimeLeft(calculateTime()), 60000);
    return () => clearInterval(interval);
  }, [targetDate]);
  return <span className="font-mono font-bold text-[10px] opacity-90">{label}: {timeLeft}</span>;
};

// --- KOMPONEN UTAMA ---

export default function OptikManager() {
  // State Data Global
  const [branches, setBranches] = useState(INITIAL_BRANCHES);
  const [technicians, setTechnicians] = useState(INITIAL_TECHNICIANS);
  const [inventory, setInventory] = useState(INITIAL_INVENTORY); 
  const [branchStocks, setBranchStocks] = useState(INITIAL_BRANCH_STOCK); 
  const [orders, setOrders] = useState(INITIAL_ORDERS);
  const [stockRequests, setStockRequests] = useState([]);
  
  // App Configuration & Subscription State
  const [appConfig, setAppConfig] = useState(INITIAL_APP_CONFIG);
  
  const [inventoryLogs, setInventoryLogs] = useState([
    { id: 1, date: '2023-10-24 08:00', item: 'Frame Rayban Aviator', type: 'IN', qty: 100, note: 'Stok Awal Pusat' },
  ]);

  // State App Logic
  const [activeRole, setActiveRole] = useState('POS'); 
  const [currentBranch, setCurrentBranch] = useState(INITIAL_BRANCHES[0]);
  const [cart, setCart] = useState([]);

  // POS State
  const [customerName, setCustomerName] = useState('');
  const [prescription, setPrescription] = useState({
    od: { sph: '', cyl: '', axis: '', add: '' },
    os: { sph: '', cyl: '', axis: '', add: '' },
    pd: ''
  });
  const [selectedProductionHub, setSelectedProductionHub] = useState('B4');

  // Payment & Cashier State
  const [showPaymentModal, setShowPaymentModal] = useState(false);
  const [showReceipt, setShowReceipt] = useState(false);
  const [paymentMethod, setPaymentMethod] = useState('CASH'); // CASH, DEBIT, QRIS
  const [cashGiven, setCashGiven] = useState('');
  const [lastOrder, setLastOrder] = useState(null);

  // ADMIN & SETTINGS State
  const [isAdminLoggedIn, setIsAdminLoggedIn] = useState(false);
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [adminPasswordInput, setAdminPasswordInput] = useState('');
  const [adminTab, setAdminTab] = useState('BRANCHES'); // BRANCHES, INVENTORY, TECHS, SETTINGS
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [settingsTab, setSettingsTab] = useState('GENERAL'); // GENERAL, LICENSE, INTEGRATION
  const [tempGSheetUrl, setTempGSheetUrl] = useState('');

  // Admin Form States
  const [newBranch, setNewBranch] = useState({ name: '', type: 'SALES' });
  const [newItem, setNewItem] = useState({ name: '', type: 'frame', stock: 0, price: 0 });
  const [newTech, setNewTech] = useState('');

  // Warehouse Modal States
  const [showRestockModal, setShowRestockModal] = useState(false);
  const [showOpnameModal, setShowOpnameModal] = useState(false);
  const [showRequestModal, setShowRequestModal] = useState(false);
  const [selectedInventoryItem, setSelectedInventoryItem] = useState('');
  const [inventoryActionQty, setInventoryActionQty] = useState('');
  const [inventoryActionNote, setInventoryActionNote] = useState('');

  // --- HELPERS UTILS ---
  const getBranchQty = (branchId, itemId) => {
    return branchStocks[branchId]?.[itemId] || 0;
  };

  // Trial & Subscription Checks (Logic from Reference)
  const isTrialExpired = useMemo(() => {
    if (!appConfig.trialEnabled || !appConfig.trialEndDate) return false;
    return new Date() > new Date(appConfig.trialEndDate);
  }, [appConfig.trialEnabled, appConfig.trialEndDate]);

  const isSubscriptionExpired = useMemo(() => {
    if (!appConfig.subscriptionEnabled || !appConfig.subscriptionEndDate) return false;
    return new Date() > new Date(appConfig.subscriptionEndDate);
  }, [appConfig.subscriptionEnabled, appConfig.subscriptionEndDate]);

  const isAccessBlocked = useMemo(() => {
    // If subscription is active and valid, allow access regardless of trial
    if (appConfig.subscriptionEnabled && !isSubscriptionExpired) return false;
    
    // If trial is enabled, check if expired
    if (appConfig.trialEnabled) return isTrialExpired;
    
    // If neither is enabled/valid
    return true; 
  }, [appConfig.trialEnabled, isTrialExpired, appConfig.subscriptionEnabled, isSubscriptionExpired]);

  // --- LOGIC AUTH ADMIN ---
  const handleRoleChange = (e) => {
    const val = e.target.value;
    
    // Check Subscription Block
    if (isAccessBlocked && val !== 'ADMIN') {
      alert("Akses Diblokir: Masa aktif aplikasi telah habis atau belum diaktifkan. Silakan hubungi Admin.");
      return;
    }

    if (val === 'ADMIN') {
      if (isAdminLoggedIn) {
        setActiveRole('ADMIN');
        setCurrentBranch({ id: 'ADMIN', name: 'Administrator' });
      } else {
        setShowAdminLogin(true);
      }
    } else if (val === 'GUDANG') {
      setActiveRole('WAREHOUSE');
      setCurrentBranch({ id: 'GUDANG', name: 'Gudang Pusat' });
    } else {
      const branch = branches.find(b => b.id === val);
      setCurrentBranch(branch);
      if (branch.type === 'PRODUCTION') setActiveRole('PRODUCTION');
      else setActiveRole('POS');
    }
  };

  const attemptAdminLogin = () => {
    if (adminPasswordInput === 'Adminoptik123!') {
      setIsAdminLoggedIn(true);
      setShowAdminLogin(false);
      setActiveRole('ADMIN');
      setCurrentBranch({ id: 'ADMIN', name: 'Administrator' });
      setAdminPasswordInput('');
    } else {
      alert('Password Salah!');
    }
  };

  const handleLogoutAdmin = () => {
    setIsAdminLoggedIn(false);
    setActiveRole('POS'); 
    setCurrentBranch(INITIAL_BRANCHES[0]);
    setShowSettingsModal(false);
  };

  // --- LOGIC SETTINGS ---
  const handleConnectGSheet = () => {
    if(!tempGSheetUrl.includes('google.com/spreadsheets')) {
      alert("URL Google Sheet tidak valid! Harap masukkan URL yang benar.");
      return;
    }
    setTimeout(() => {
      setAppConfig({ ...appConfig, dbType: 'GSHEET', gsheetUrl: tempGSheetUrl });
      alert("Berhasil terhubung ke Google Sheet! Data akan disinkronisasi secara otomatis.");
    }, 1000);
  };

  const handleSwitchToLocal = () => {
    setAppConfig({ ...appConfig, dbType: 'LOCAL', gsheetUrl: '' });
    alert("Mode beralih ke Database Lokal.");
  };

  // --- LOGIC ADMIN ACTIONS ---
  const handleAddBranch = () => {
    if (!newBranch.name) return alert('Nama cabang harus diisi');
    const id = `B${branches.length + 1 + Math.floor(Math.random() * 100)}`;
    setBranches([...branches, { id, ...newBranch }]);
    setNewBranch({ name: '', type: 'SALES' });
    setBranchStocks({ ...branchStocks, [id]: {} });
    alert('Cabang berhasil ditambahkan');
  };

  const handleAddItem = () => {
    if (!newItem.name) return alert('Nama item harus diisi');
    const id = newItem.type === 'frame' ? `F${Date.now()}` : newItem.type === 'lens' ? `L${Date.now()}` : `P${Date.now()}`;
    setInventory([...inventory, { id, ...newItem }]);
    setInventoryLogs([...inventoryLogs, {
      id: Date.now(),
      date: getCurrentTime(),
      item: newItem.name,
      type: 'IN',
      qty: newItem.stock,
      note: 'Item Baru (Admin)'
    }]);
    setNewItem({ name: '', type: 'frame', stock: 0, price: 0 });
    alert('Item berhasil ditambahkan ke inventori pusat');
  };

  const handleAddTech = () => {
    if (!newTech) return alert('Nama teknisi harus diisi');
    const id = `T${Date.now()}`;
    setTechnicians([...technicians, { id, name: newTech }]);
    setNewTech('');
    alert('Teknisi berhasil ditambahkan');
  };

  // --- LOGIC WAREHOUSE & POS (Simplified for brevity as they are unchanged) ---
  const handleRestock = () => { /* ... existing logic ... */ 
    const qty = parseInt(inventoryActionQty);
    const item = inventory.find(i => i.id === selectedInventoryItem);
    if (!item || !qty) return;
    setInventory(inventory.map(i => i.id === selectedInventoryItem ? { ...i, stock: i.stock + qty } : i));
    setInventoryLogs([{ id: Date.now(), date: getCurrentTime(), item: item.name, type: 'IN', qty: qty, note: inventoryActionNote || 'Restock' }, ...inventoryLogs]);
    setShowRestockModal(false); setInventoryActionQty(''); setInventoryActionNote('');
  };
  const handleOpname = () => { /* ... existing logic ... */
    const actualQty = parseInt(inventoryActionQty);
    const item = inventory.find(i => i.id === selectedInventoryItem);
    if (!item) return;
    const diff = actualQty - item.stock;
    const type = diff >= 0 ? 'ADJUST_IN' : 'ADJUST_OUT';
    setInventory(inventory.map(i => i.id === selectedInventoryItem ? { ...i, stock: actualQty } : i));
    setInventoryLogs([{ id: Date.now(), date: getCurrentTime(), item: item.name, type: type, qty: Math.abs(diff), note: inventoryActionNote || 'Opname' }, ...inventoryLogs]);
    setShowOpnameModal(false); setInventoryActionQty(''); setInventoryActionNote('');
  };
  const handleOpenRequestModal = (item) => { setSelectedInventoryItem(item.id); setInventoryActionQty(''); setInventoryActionNote(''); setShowRequestModal(true); };
  const submitStockRequest = () => { /* ... existing logic ... */
    const qty = parseInt(inventoryActionQty);
    const item = inventory.find(i => i.id === selectedInventoryItem);
    setStockRequests([...stockRequests, { id: `REQ-${Date.now()}`, branchId: currentBranch.id, branchName: currentBranch.name, itemId: selectedInventoryItem, itemName: item.name, qty, note: inventoryActionNote, status: 'PENDING', date: getCurrentTime() }]);
    setShowRequestModal(false);
  };
  const approveStockRequest = (reqId) => { /* ... existing logic ... */
    const req = stockRequests.find(r => r.id === reqId);
    const item = inventory.find(i => i.id === req.itemId);
    if (item.stock < req.qty) return alert("Stok Pusat Kurang");
    setInventory(inventory.map(i => i.id === req.itemId ? { ...i, stock: i.stock - req.qty } : i));
    const currentBranchStock = branchStocks[req.branchId] || {};
    setBranchStocks({ ...branchStocks, [req.branchId]: { ...currentBranchStock, [req.itemId]: (currentBranchStock[req.itemId] || 0) + req.qty } });
    setStockRequests(stockRequests.map(r => r.id === reqId ? { ...r, status: 'APPROVED' } : r));
    setInventoryLogs([{ id: Date.now(), date: getCurrentTime(), item: item.name, type: 'OUT', qty: req.qty, note: `Kirim ke ${req.branchName}` }, ...inventoryLogs]);
  };
  const addToCart = (item) => setCart([...cart, item]);
  const removeFromCart = (index) => { const newCart = [...cart]; newCart.splice(index, 1); setCart(newCart); };
  const calculateTotal = () => cart.reduce((a, b) => a + b.price, 0);
  const handleCheckout = () => { if (customerName && cart.length) setShowPaymentModal(true); };
  const processPayment = () => {
    const total = calculateTotal();
    const cash = parseFloat(cashGiven) || 0;
    if (paymentMethod === 'CASH' && cash < total) { alert("Uang Kurang"); return; }
    const newBranchStock = { ...branchStocks[currentBranch.id] };
    cart.forEach(item => { if (newBranchStock[item.id]) newBranchStock[item.id] -= 1; });
    setBranchStocks({ ...branchStocks, [currentBranch.id]: newBranchStock });
    setOrders([{ id: `ORD-${Math.floor(Math.random()*10000)}`, customer: customerName, branchId: currentBranch.id, productionBranchId: selectedProductionHub, status: 'PENDING', technician: null, items: [...cart], totalPrice: total, paymentMethod, prescription, logs: [{date: getCurrentTime(), msg: 'Order Masuk', type: 'info'}], partsUsed: [] }, ...orders]);
    setLastOrder({ id: 'NEW', customer: customerName, items: [...cart], totalPrice: total, cashGiven: cash, change: cash - total, paymentMethod, logs: [{date: getCurrentTime()}] });
    setCart([]); setCustomerName(''); setPrescription({ od: { sph: '', cyl: '', axis: '', add: '' }, os: { sph: '', cyl: '', axis: '', add: '' }, pd: '' });
    setShowPaymentModal(false); setShowReceipt(true); setCashGiven('');
  };
  const assignTechnicianAndStart = (orderId, techName) => {
    if (!techName) { alert("Pilih Teknisi"); return; }
    setOrders(orders.map(o => o.id === orderId ? { ...o, status: 'IN_PROGRESS', technician: techName, logs: [...o.logs, { date: getCurrentTime(), msg: `Dikerjakan: ${techName}`, type: 'progress' }] } : o));
  };
  const updateOrderStatus = (orderId, newStatus) => {
    setOrders(orders.map(o => o.id === orderId ? { ...o, status: newStatus, logs: [...o.logs, { date: getCurrentTime(), msg: `Status: ${newStatus}`, type: 'progress' }] } : o));
  };
  const useSparePart = (orderId, partId, isDefect = false) => {
    const part = inventory.find(i => i.id === partId);
    const localQty = getBranchQty(currentBranch.id, partId);
    if (localQty <= 0) { alert(`Stok ${part.name} Habis`); return; }
    setOrders(orders.map(o => o.id === orderId ? { ...o, partsUsed: [...o.partsUsed, { partId, name: part.name, status: isDefect ? 'BROKEN' : 'USED' }], logs: [...o.logs, { date: getCurrentTime(), msg: `${isDefect ? 'RUSAK' : 'PAKAI'}: ${part.name}`, type: isDefect ? 'danger' : 'part' }] } : o));
    const newLocalStock = { ...branchStocks[currentBranch.id] }; newLocalStock[partId] = localQty - 1;
    setBranchStocks({ ...branchStocks, [currentBranch.id]: newLocalStock });
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 relative">
      
      {/* --- BANNER SUBSCRIPTION / DB STATUS --- */}
      <div className="bg-slate-900 text-white text-xs py-1 px-4 flex flex-col md:flex-row justify-between items-center gap-2 md:gap-0">
        <div className="flex items-center gap-4">
          <span className="flex items-center gap-1 opacity-80">
            {appConfig.dbType === 'GSHEET' ? <FileSpreadsheet size={12} className="text-green-400" /> : <Database size={12} className="text-blue-400" />}
            <span className="hidden md:inline">Mode:</span> {appConfig.dbType === 'GSHEET' ? 'Google Sheet Sync' : 'Database Lokal'}
          </span>
        </div>
        <div>
          {appConfig.subscriptionEnabled && !isSubscriptionExpired ? (
             <span className="flex items-center gap-1 text-teal-400 font-bold">
               <Crown size={12} /> PREMIUM LICENSE
               <span className="bg-white/10 px-2 py-0.5 rounded text-[9px] ml-1">
                 <CountdownDisplay targetDate={appConfig.subscriptionEndDate} label="Sisa" />
               </span>
             </span>
          ) : appConfig.trialEnabled && !isTrialExpired ? (
            <span className="text-yellow-400 font-bold flex items-center gap-2">
              TRIAL MODE 
              <span className="bg-white/10 px-2 py-0.5 rounded text-[9px]">
                <CountdownDisplay targetDate={appConfig.trialEndDate} label="Sisa" />
              </span>
            </span>
          ) : (
            <span className="text-red-400 font-bold flex items-center gap-1"><Lock size={12}/> EXPIRED</span>
          )}
        </div>
      </div>

      {/* --- BLOCKED ACCESS OVERLAY --- */}
      {isAccessBlocked && !isAdminLoggedIn && (
        <div className="fixed inset-0 bg-slate-900/95 z-[999] flex flex-col items-center justify-center text-center p-8">
          <AlertTriangle size={64} className="text-red-500 mb-4" />
          <h2 className="text-3xl font-bold text-white mb-2">Masa Aktif Berakhir</h2>
          <p className="text-slate-400 max-w-md mb-8">Akses ke aplikasi dibatasi karena masa trial atau langganan telah habis. Silakan hubungi Administrator atau login sebagai Admin untuk memperbarui pengaturan.</p>
          <button 
            onClick={() => setShowAdminLogin(true)} 
            className="bg-red-600 hover:bg-red-700 text-white px-6 py-3 rounded-lg font-bold transition flex items-center gap-2"
          >
            <Lock size={18} /> Login Admin
          </button>
        </div>
      )}

      {/* --- MODAL PENGATURAN (SETTINGS) --- */}
      {showSettingsModal && (
        <div className="fixed inset-0 bg-black/50 z-[70] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl overflow-hidden flex flex-col h-auto max-h-[90vh]">
            <div className="bg-slate-800 text-white p-4 flex justify-between items-center shrink-0">
              <h3 className="font-bold text-lg flex items-center gap-2">
                <Settings size={20} /> Pengaturan Sistem
              </h3>
              <button onClick={() => setShowSettingsModal(false)} className="hover:bg-slate-700 p-1 rounded">
                <X size={20} />
              </button>
            </div>
            
            <div className="flex flex-col md:flex-row flex-1 overflow-hidden">
              {/* SIDEBAR SETTINGS - MOBILE SCROLLABLE TAB, DESKTOP SIDEBAR */}
              <div className="w-full md:w-48 bg-slate-100 border-b md:border-b-0 md:border-r p-2 md:p-4 flex md:flex-col gap-2 overflow-x-auto md:overflow-y-auto shrink-0">
                <button 
                  onClick={() => setSettingsTab('GENERAL')}
                  className={`flex-1 md:flex-none text-center md:text-left px-3 py-2 rounded text-sm font-medium whitespace-nowrap ${settingsTab === 'GENERAL' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-600 hover:bg-slate-200'}`}
                >
                  Umum & Database
                </button>
                <button 
                  onClick={() => setSettingsTab('LICENSE')}
                  className={`flex-1 md:flex-none text-center md:text-left px-3 py-2 rounded text-sm font-medium whitespace-nowrap ${settingsTab === 'LICENSE' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-600 hover:bg-slate-200'}`}
                >
                  Lisensi & Trial
                </button>
                <button 
                  onClick={() => setSettingsTab('INTEGRATION')}
                  className={`flex-1 md:flex-none text-center md:text-left px-3 py-2 rounded text-sm font-medium whitespace-nowrap ${settingsTab === 'INTEGRATION' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-600 hover:bg-slate-200'}`}
                >
                  Integrasi Script
                </button>
              </div>

              {/* CONTENT SETTINGS */}
              <div className="flex-1 p-4 md:p-6 overflow-y-auto min-h-0">
                
                {/* TAB: GENERAL */}
                {settingsTab === 'GENERAL' && (
                  <div className="space-y-6">
                    <h4 className="font-bold text-slate-800 text-lg border-b pb-2">Koneksi Database</h4>
                    <div className="space-y-4">
                      <div 
                        onClick={handleSwitchToLocal}
                        className={`border p-4 rounded-lg cursor-pointer flex items-center gap-3 transition ${appConfig.dbType === 'LOCAL' ? 'border-blue-500 bg-blue-50' : 'hover:bg-slate-50'}`}
                      >
                        <Database size={24} className="text-slate-500" />
                        <div>
                          <p className="font-bold text-sm">Database Lokal (Offline)</p>
                          <p className="text-xs text-slate-500">Penyimpanan di browser perangkat ini.</p>
                        </div>
                        {appConfig.dbType === 'LOCAL' && <CheckCircle size={16} className="text-blue-600 ml-auto" />}
                      </div>

                      <div className={`border p-4 rounded-lg transition ${appConfig.dbType === 'GSHEET' ? 'border-green-500 bg-green-50' : ''}`}>
                        <div className="flex items-center gap-3 mb-3">
                          <FileSpreadsheet size={24} className="text-green-600" />
                          <div>
                            <p className="font-bold text-sm">Google Sheets (Online)</p>
                            <p className="text-xs text-slate-500">Sinkronisasi data real-time via API.</p>
                          </div>
                          {appConfig.dbType === 'GSHEET' && <CheckCircle size={16} className="text-green-600 ml-auto" />}
                        </div>
                        <div className="flex flex-col md:flex-row gap-2">
                          <input 
                            type="text" 
                            className="flex-1 text-xs p-3 border rounded-lg" 
                            placeholder="Tempel URL Web App Google Script..."
                            value={tempGSheetUrl || appConfig.gsheetUrl}
                            onChange={(e) => setTempGSheetUrl(e.target.value)}
                          />
                          <button onClick={handleConnectGSheet} className="bg-green-600 text-white text-xs px-4 py-2 rounded-lg font-bold hover:bg-green-700">Hubungkan</button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* TAB: LICENSE (REFERENCED STYLE) */}
                {settingsTab === 'LICENSE' && (
                  <div className="space-y-6">
                    <h4 className="font-bold text-slate-800 text-lg border-b pb-2 flex items-center gap-2">
                      <Crown size={20} className="text-purple-600" /> Manajemen Masa Aktif
                    </h4>
                    
                    <div className="bg-yellow-50 border border-yellow-200 p-3 rounded text-xs text-yellow-800 mb-4">
                      <strong>Info Admin:</strong> Anda memiliki kontrol penuh untuk mengubah status dan masa aktif aplikasi ini.
                    </div>

                    {/* SECTION: TRIAL */}
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 space-y-6">
                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-3">
                            <div 
                              className="w-10 h-10 rounded-xl flex items-center justify-center text-white" 
                              style={{ backgroundColor: appConfig.trialEnabled ? '#f59e0b' : '#cbd5e1' }}
                            >
                              <Timer className="w-5 h-5" />
                            </div>
                            <div>
                              <h3 className="text-sm font-black text-slate-900 uppercase">Mode Trial</h3>
                              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                                Status: {appConfig.trialEnabled ? 'AKTIF' : 'NONAKTIF'}
                              </p>
                            </div>
                          </div>
                          <button 
                            onClick={() => setAppConfig(prev => ({ 
                              ...prev, 
                              trialEnabled: !prev.trialEnabled, 
                              subscriptionEnabled: false // Disable subs if trial enabled to avoid conflict logic visually
                            }))} 
                            className={`w-14 h-8 rounded-full p-1 transition-all ${appConfig.trialEnabled ? 'bg-amber-500' : 'bg-slate-200'}`}
                          >
                            <div className={`w-6 h-6 bg-white rounded-full shadow-sm transform transition-all ${appConfig.trialEnabled ? 'translate-x-6' : ''}`}></div>
                          </button>
                        </div>
                        
                        {appConfig.trialEnabled && (
                          <div className="space-y-4 pt-4 border-t border-slate-50 animate-in slide-in-from-top-2">
                            <div>
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block ml-2 mb-2">
                                Tanggal Berakhir
                              </label>
                              <input 
                                type="datetime-local" 
                                className="w-full p-4 border-2 border-slate-100 rounded-2xl text-xs font-bold text-slate-700 outline-none focus:border-amber-500 mb-3" 
                                value={appConfig.trialEndDate} 
                                onChange={(e) => setAppConfig({ ...appConfig, trialEndDate: e.target.value })} 
                              />
                              <div className="flex gap-2">
                                <button 
                                  onClick={() => setAppConfig({ ...appConfig, trialEndDate: addTime(appConfig.trialEndDate, 'days', 7) })} 
                                  className="flex-1 py-3 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-bold text-slate-500 hover:bg-slate-100"
                                >
                                  +7 Hari
                                </button>
                                <button 
                                  onClick={() => setAppConfig({ ...appConfig, trialEndDate: addTime(appConfig.trialEndDate, 'days', 30) })} 
                                  className="flex-1 py-3 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-bold text-slate-500 hover:bg-slate-100"
                                >
                                  +30 Hari
                                </button>
                              </div>
                            </div>
                            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 text-center">
                              <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">
                                Sisa Waktu
                              </p>
                              <div className="text-amber-600">
                                <CountdownDisplay targetDate={appConfig.trialEndDate} label="Trial" />
                              </div>
                            </div>
                          </div>
                        )}
                    </div>

                    {/* SECTION: SUBSCRIPTION */}
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 space-y-6">
                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-3">
                            <div 
                              className="w-10 h-10 rounded-xl flex items-center justify-center text-white" 
                              style={{ backgroundColor: appConfig.subscriptionEnabled ? '#2563eb' : '#cbd5e1' }}
                            >
                              <CreditCard className="w-5 h-5" />
                            </div>
                            <div>
                              <h3 className="text-sm font-black text-slate-900 uppercase">Berlangganan</h3>
                              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                                Status: {appConfig.subscriptionEnabled ? 'AKTIF' : 'NONAKTIF'}
                              </p>
                            </div>
                          </div>
                          <button 
                            onClick={() => setAppConfig(prev => ({ 
                              ...prev, 
                              subscriptionEnabled: !prev.subscriptionEnabled, 
                              trialEnabled: false // Disable trial if subs enabled
                            }))} 
                            className={`w-14 h-8 rounded-full p-1 transition-all ${appConfig.subscriptionEnabled ? 'bg-blue-500' : 'bg-slate-200'}`}
                          >
                            <div className={`w-6 h-6 bg-white rounded-full shadow-sm transform transition-all ${appConfig.subscriptionEnabled ? 'translate-x-6' : ''}`}></div>
                          </button>
                        </div>
                        
                        {appConfig.subscriptionEnabled && (
                          <div className="space-y-4 pt-4 border-t border-slate-50 animate-in slide-in-from-top-2">
                            <div>
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block ml-2 mb-2">
                                Tanggal Berakhir
                              </label>
                              <input 
                                type="datetime-local" 
                                className="w-full p-4 border-2 border-slate-100 rounded-2xl text-xs font-bold text-slate-700 outline-none focus:border-blue-500 mb-3" 
                                value={appConfig.subscriptionEndDate} 
                                onChange={(e) => setAppConfig({ ...appConfig, subscriptionEndDate: e.target.value })} 
                              />
                              <div className="flex flex-wrap gap-2 mb-2">
                                <button 
                                  onClick={() => setAppConfig({ ...appConfig, subscriptionEndDate: addTime(appConfig.subscriptionEndDate, 'months', 1) })} 
                                  className="flex-1 py-3 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-bold text-slate-500 hover:bg-slate-100"
                                >
                                  +1 Bulan
                                </button>
                                <button 
                                  onClick={() => setAppConfig({ ...appConfig, subscriptionEndDate: addTime(appConfig.subscriptionEndDate, 'months', 6) })} 
                                  className="flex-1 py-3 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-bold text-slate-500 hover:bg-slate-100"
                                >
                                  +6 Bulan
                                </button>
                                <button 
                                  onClick={() => setAppConfig({ ...appConfig, subscriptionEndDate: addTime(appConfig.subscriptionEndDate, 'years', 1) })} 
                                  className="flex-1 py-3 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-bold text-slate-500 hover:bg-slate-100"
                                >
                                  +1 Tahun
                                </button>
                              </div>
                            </div>
                            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 text-center">
                              <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">
                                Sisa Waktu
                              </p>
                              <div className="text-blue-600">
                                <CountdownDisplay targetDate={appConfig.subscriptionEndDate} label="Langganan" />
                              </div>
                            </div>
                          </div>
                        )}
                    </div>

                  </div>
                )}

                {/* TAB: INTEGRATION (GAS SCRIPT) */}
                {settingsTab === 'INTEGRATION' && (
                  <div className="space-y-4 h-full flex flex-col">
                    <h4 className="font-bold text-slate-800 text-lg border-b pb-2 flex items-center gap-2">
                      <Code size={20} className="text-slate-600" /> Google Apps Script (Backend)
                    </h4>
                    
                    <p className="text-xs text-slate-500">
                      Untuk menghubungkan aplikasi ini ke Google Sheet, buat Spreadhseet baru, buka <b>Extensions {'>'} Apps Script</b>, lalu tempel kode berikut:
                    </p>

                    <div className="relative flex-1 min-h-[200px] bg-slate-900 rounded-lg overflow-hidden border border-slate-700">
                      <textarea 
                        readOnly 
                        className="w-full h-full bg-slate-900 text-green-400 font-mono text-xs p-4 resize-none focus:outline-none"
                        value={GAS_TEMPLATE}
                      />
                      <button 
                        onClick={() => {navigator.clipboard.writeText(GAS_TEMPLATE); alert("Kode disalin ke clipboard!");}}
                        className="absolute top-2 right-2 bg-white/10 hover:bg-white/20 text-white px-3 py-1 rounded text-xs backdrop-blur-sm border border-white/20"
                      >
                        Salin Kode
                      </button>
                    </div>
                    
                    <div className="text-xs text-slate-500 bg-blue-50 p-2 rounded text-center border border-blue-100">
                      Setelah save, lakukan <b>Deploy {'>'} New Deployment</b>. Pilih type "Web App", Execute as "Me", dan Who has access "Anyone". Salin URL Web App ke tab <b>Umum & Database</b>.
                    </div>
                  </div>
                )}

              </div>
            </div>
            
            <div className="p-4 bg-slate-50 border-t flex justify-between items-center shrink-0">
              <span className="text-[10px] text-slate-400 italic hidden md:block">Perubahan disimpan otomatis di sesi ini.</span>
              <div className="flex gap-2 w-full md:w-auto">
                <button 
                  onClick={() => setShowSettingsModal(false)}
                  className="flex-1 md:flex-none bg-slate-200 text-slate-700 px-4 py-3 rounded-lg font-bold text-sm hover:bg-slate-300"
                >
                  Tutup
                </button>
                {isAdminLoggedIn && (
                  <button 
                    onClick={handleLogoutAdmin}
                    className="flex-1 md:flex-none bg-red-600 text-white px-4 py-3 rounded-lg font-bold text-sm hover:bg-red-700"
                  >
                    Logout Admin
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* --- MODAL LOGIN ADMIN --- */}
      {showAdminLogin && (
        <div className="fixed inset-0 bg-slate-900/80 z-[60] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm overflow-hidden p-6">
            <div className="flex flex-col items-center mb-6">
              <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mb-4 text-red-600">
                <Lock size={32} />
              </div>
              <h3 className="text-xl font-bold text-slate-800">Akses Administrator</h3>
              <p className="text-sm text-slate-500 text-center mt-2">Masukkan password untuk mengelola sistem.</p>
            </div>
            <input 
              type="password" 
              className="w-full p-3 border rounded-lg mb-4 focus:ring-2 focus:ring-red-500 outline-none text-center tracking-widest"
              placeholder="Password..."
              value={adminPasswordInput}
              onChange={(e) => setAdminPasswordInput(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && attemptAdminLogin()}
            />
            <div className="flex gap-2">
              <button onClick={() => setShowAdminLogin(false)} className="flex-1 py-2 border rounded-lg hover:bg-slate-50 font-semibold text-slate-600">Batal</button>
              <button onClick={attemptAdminLogin} className="flex-1 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg font-semibold">Login</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL REQUEST STOCK, RESTOCK, OPNAME (SAMA SEPERTI SEBELUMNYA) */}
      {showRequestModal && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6">
            <h3 className="font-bold text-lg mb-4 flex items-center gap-2 text-teal-700"><Send size={24} /> Request ke Gudang Pusat</h3>
            <p className="text-xs text-slate-500 mb-4">Permintaan dari: <span className="font-bold text-slate-700">{currentBranch.name}</span></p>
            <div className="space-y-4">
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Item</label><input type="text" className="w-full border p-2 rounded bg-slate-100" value={inventory.find(i => i.id === selectedInventoryItem)?.name || ''} readOnly /></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Qty</label><input type="number" className="w-full border p-2 rounded" placeholder="0" value={inventoryActionQty} onChange={e => setInventoryActionQty(e.target.value)} autoFocus /></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Catatan</label><input type="text" className="w-full border p-2 rounded" value={inventoryActionNote} onChange={e => setInventoryActionNote(e.target.value)} /></div>
            </div>
            <div className="flex gap-2 mt-6"><button onClick={() => setShowRequestModal(false)} className="flex-1 border py-2 text-sm hover:bg-slate-50">Batal</button><button onClick={submitStockRequest} className="flex-1 bg-teal-600 text-white py-2 text-sm font-bold hover:bg-teal-700">Kirim</button></div>
          </div>
        </div>
      )}
      {showRestockModal && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6">
            <h3 className="font-bold text-lg mb-4 flex items-center gap-2 text-teal-700"><PackagePlus size={24} /> Restock Gudang Pusat</h3>
            <div className="space-y-4">
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Item</label><select className="w-full border p-2 rounded" value={selectedInventoryItem} onChange={e => setSelectedInventoryItem(e.target.value)}><option value="">-- Pilih --</option>{inventory.map(i => (<option key={i.id} value={i.id}>{i.name}</option>))}</select></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Qty</label><input type="number" className="w-full border p-2 rounded" placeholder="0" value={inventoryActionQty} onChange={e => setInventoryActionQty(e.target.value)} /></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Sumber</label><input type="text" className="w-full border p-2 rounded" value={inventoryActionNote} onChange={e => setInventoryActionNote(e.target.value)} /></div>
            </div>
            <div className="flex gap-2 mt-6"><button onClick={() => setShowRestockModal(false)} className="flex-1 border py-2 text-sm hover:bg-slate-50">Batal</button><button onClick={handleRestock} className="flex-1 bg-teal-600 text-white py-2 text-sm font-bold hover:bg-teal-700">Simpan</button></div>
          </div>
        </div>
      )}
      {showOpnameModal && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6">
            <h3 className="font-bold text-lg mb-4 flex items-center gap-2 text-orange-600"><ClipboardCheck size={24} /> Stock Opname Pusat</h3>
            <div className="space-y-4">
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Item</label><select className="w-full border p-2 rounded" value={selectedInventoryItem} onChange={e => setSelectedInventoryItem(e.target.value)}><option value="">-- Pilih --</option>{inventory.map(i => (<option key={i.id} value={i.id}>{i.name} (Sistem: {i.stock})</option>))}</select></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Fisik Aktual</label><input type="number" className="w-full border p-2 rounded font-bold text-orange-700" placeholder="0" value={inventoryActionQty} onChange={e => setInventoryActionQty(e.target.value)} /></div>
              <div><label className="text-xs font-bold text-slate-500 block mb-1">Alasan</label><input type="text" className="w-full border p-2 rounded" value={inventoryActionNote} onChange={e => setInventoryActionNote(e.target.value)} /></div>
            </div>
            <div className="flex gap-2 mt-6"><button onClick={() => setShowOpnameModal(false)} className="flex-1 border py-2 text-sm hover:bg-slate-50">Batal</button><button onClick={handleOpname} className="flex-1 bg-orange-600 text-white py-2 text-sm font-bold hover:bg-orange-700">Simpan</button></div>
          </div>
        </div>
      )}
      {showPaymentModal && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden">
            <div className="bg-teal-700 text-white p-4 flex justify-between items-center"><h3 className="font-bold text-lg flex items-center gap-2"><Banknote size={20} /> Kasir / Pembayaran</h3><button onClick={() => setShowPaymentModal(false)} className="hover:bg-teal-600 p-1 rounded"><X size={20} /></button></div>
            <div className="p-6">
              <div className="mb-6 text-center"><p className="text-sm text-slate-500 mb-1">Total Tagihan</p><div className="text-3xl font-bold text-slate-800">{formatRupiah(calculateTotal())}</div></div>
              <div className="mb-4"><label className="block text-sm font-semibold mb-2">Metode Pembayaran</label><div className="grid grid-cols-3 gap-2"><button onClick={() => setPaymentMethod('CASH')} className={`p-3 rounded border flex flex-col items-center gap-1 text-xs font-bold ${paymentMethod === 'CASH' ? 'bg-teal-50 border-teal-500 text-teal-700' : 'hover:bg-slate-50'}`}><Banknote size={20} /> TUNAI</button><button onClick={() => setPaymentMethod('DEBIT')} className={`p-3 rounded border flex flex-col items-center gap-1 text-xs font-bold ${paymentMethod === 'DEBIT' ? 'bg-teal-50 border-teal-500 text-teal-700' : 'hover:bg-slate-50'}`}><CreditCard size={20} /> DEBIT/CC</button><button onClick={() => setPaymentMethod('QRIS')} className={`p-3 rounded border flex flex-col items-center gap-1 text-xs font-bold ${paymentMethod === 'QRIS' ? 'bg-teal-50 border-teal-500 text-teal-700' : 'hover:bg-slate-50'}`}><QrCode size={20} /> QRIS</button></div></div>
              {paymentMethod === 'CASH' && (<div className="mb-6"><label className="block text-sm font-semibold mb-2">Uang Diterima</label><input type="number" className="w-full p-3 text-lg border rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" placeholder="Masukkan jumlah uang..." value={cashGiven} onChange={(e) => setCashGiven(e.target.value)} />{cashGiven && (<div className="mt-2 text-right"><span className="text-sm text-slate-500">Kembalian: </span><span className={`font-bold ${parseFloat(cashGiven) < calculateTotal() ? 'text-red-500' : 'text-green-600'}`}>{formatRupiah(parseFloat(cashGiven) - calculateTotal())}</span></div>)}</div>)}
              <button onClick={processPayment} className="w-full bg-teal-600 hover:bg-teal-700 text-white py-3 rounded-lg font-bold shadow-lg shadow-teal-200 transition">Bayar Sekarang</button>
            </div>
          </div>
        </div>
      )}
      {showReceipt && lastOrder && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm overflow-hidden relative">
            <div className="p-8 bg-white border-b-2 border-dashed border-slate-300"><div className="text-center mb-6"><div className="inline-flex items-center justify-center w-12 h-12 bg-teal-100 text-teal-700 rounded-full mb-2"><CheckCircle size={24} /></div><h3 className="font-bold text-xl text-slate-800">Pembayaran Berhasil!</h3><p className="text-xs text-slate-500">{lastOrder.logs[0].date}</p></div><div className="space-y-4 text-sm"><div className="flex justify-between border-b pb-2"><span className="text-slate-500">No. Order</span><span className="font-mono font-bold">{lastOrder.id}</span></div><div className="flex justify-between border-b pb-2"><span className="text-slate-500">Pelanggan</span><span className="font-bold">{lastOrder.customer}</span></div><div className="py-2"><p className="text-xs text-slate-400 mb-1 uppercase tracking-wider">Items</p>{lastOrder.items.map((item, i) => (<div key={i} className="flex justify-between mb-1"><span>{item.name}</span><span>{formatRupiah(item.price)}</span></div>))}</div><div className="border-t pt-2 space-y-1"><div className="flex justify-between font-bold text-lg"><span>Total</span><span>{formatRupiah(lastOrder.totalPrice)}</span></div>{lastOrder.paymentMethod === 'CASH' && (<><div className="flex justify-between text-slate-500"><span>Tunai</span><span>{formatRupiah(lastOrder.cashGiven)}</span></div><div className="flex justify-between text-teal-600 font-bold"><span>Kembalian</span><span>{formatRupiah(lastOrder.change)}</span></div></>)}</div></div></div>
            <div className="p-4 bg-slate-50 flex gap-2"><button onClick={() => window.print()} className="flex-1 border border-slate-300 text-slate-700 py-2 rounded-lg font-bold text-sm hover:bg-white transition flex items-center justify-center gap-2"><Receipt size={16} /> Cetak</button><button onClick={() => setShowReceipt(false)} className="flex-1 bg-teal-600 text-white py-2 rounded-lg font-bold text-sm hover:bg-teal-700 transition">Tutup</button></div>
          </div>
        </div>
      )}

      {/* HEADER NAVIGASI ROLE */}
      <header className={`text-white shadow-lg sticky top-0 z-40 transition-colors duration-300 ${activeRole === 'ADMIN' ? 'bg-slate-800' : 'bg-teal-700'}`}>
        <div className="max-w-7xl mx-auto px-4 py-3 flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-2 w-full md:w-auto justify-between md:justify-start">
            <div className="flex items-center gap-2">
              <Glasses className={`h-8 w-8 ${activeRole === 'ADMIN' ? 'text-red-400' : 'text-teal-200'}`} />
              <div>
                <h1 className="text-xl font-bold tracking-tight">OptikVision Pro</h1>
                <p className={`text-xs opacity-90 ${activeRole === 'ADMIN' ? 'text-red-200' : 'text-teal-200'}`}>
                  {activeRole === 'ADMIN' ? 'ADMIN MODE' : 'Sistem Informasi Optik Terpadu'}
                </p>
              </div>
            </div>
            {/* Mobile Menu Toggle could be added here if needed */}
          </div>

          <div className={`flex flex-col md:flex-row items-center gap-2 md:gap-4 p-2 rounded-lg w-full md:w-auto ${activeRole === 'ADMIN' ? 'bg-slate-700' : 'bg-teal-800'}`}>
            <div className="flex flex-col w-full md:w-auto">
              <label className={`text-[10px] uppercase font-bold hidden md:block ${activeRole === 'ADMIN' ? 'text-slate-400' : 'text-teal-300'}`}>Login Sebagai:</label>
              <select 
                value={activeRole === 'ADMIN' ? 'ADMIN' : (currentBranch.id === 'GUDANG' ? 'GUDANG' : currentBranch.id)} 
                onChange={handleRoleChange}
                className={`border-none text-white text-sm rounded focus:ring-0 cursor-pointer w-full md:w-auto ${activeRole === 'ADMIN' ? 'bg-slate-900' : 'bg-teal-900'}`}
              >
                <optgroup label="Cabang & Produksi">
                  {branches.map(b => (
                    <option key={b.id} value={b.id}>{b.name}</option>
                  ))}
                </optgroup>
                <option value="GUDANG">GUDANG PUSAT (Logistik)</option>
                <option value="ADMIN">ADMINISTRATOR (Pemilik)</option>
              </select>
            </div>
            
            <div className={`h-8 w-[1px] mx-2 hidden md:block ${activeRole === 'ADMIN' ? 'bg-slate-500' : 'bg-teal-600'}`}></div>

            <nav className="flex gap-2 w-full md:w-auto overflow-x-auto pb-1 md:pb-0 no-scrollbar">
              {activeRole === 'ADMIN' ? (
                <div className="px-3 py-1 bg-red-500/20 border border-red-500/50 text-red-200 rounded text-sm font-bold flex items-center gap-2 w-full justify-center md:w-auto">
                  <Lock size={14} /> Full Access
                </div>
              ) : (
                <>
                  <button onClick={() => setActiveRole('POS')} className={`flex-1 md:flex-none px-3 py-2 md:py-1 rounded text-sm font-medium whitespace-nowrap ${activeRole === 'POS' ? 'bg-white text-teal-800' : 'hover:bg-teal-600'}`}>POS</button>
                  {(currentBranch.type === 'PRODUCTION' || currentBranch.id === 'GUDANG') && (
                    <button onClick={() => setActiveRole('PRODUCTION')} className={`flex-1 md:flex-none px-3 py-2 md:py-1 rounded text-sm font-medium whitespace-nowrap ${activeRole === 'PRODUCTION' ? 'bg-white text-teal-800' : 'hover:bg-teal-600'}`}>Produksi</button>
                  )}
                  <button onClick={() => setActiveRole('WAREHOUSE')} className={`flex-1 md:flex-none px-3 py-2 md:py-1 rounded text-sm font-medium whitespace-nowrap ${activeRole === 'WAREHOUSE' ? 'bg-white text-teal-800' : 'hover:bg-teal-600'}`}>Gudang</button>
                </>
              )}
            </nav>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-4 md:p-6 pb-24 md:pb-6">

        {/* --- VIEW: ADMIN DASHBOARD --- */}
        {activeRole === 'ADMIN' && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 md:gap-4 mb-6">
              <button onClick={() => setAdminTab('BRANCHES')} className={`p-4 rounded-xl border flex flex-col md:flex-row items-center md:items-start gap-3 transition text-center md:text-left ${adminTab === 'BRANCHES' ? 'bg-slate-800 text-white shadow-lg' : 'bg-white hover:bg-slate-50'}`}>
                <div className={`p-2 rounded-full ${adminTab === 'BRANCHES' ? 'bg-slate-700' : 'bg-slate-100 text-slate-600'}`}><Building size={20} /></div>
                <div><p className="font-bold text-sm">Cabang</p><p className="text-xs opacity-70 hidden md:block">{branches.length} Aktif</p></div>
              </button>
              <button onClick={() => setAdminTab('INVENTORY')} className={`p-4 rounded-xl border flex flex-col md:flex-row items-center md:items-start gap-3 transition text-center md:text-left ${adminTab === 'INVENTORY' ? 'bg-slate-800 text-white shadow-lg' : 'bg-white hover:bg-slate-50'}`}>
                <div className={`p-2 rounded-full ${adminTab === 'INVENTORY' ? 'bg-slate-700' : 'bg-slate-100 text-slate-600'}`}><Package size={20} /></div>
                <div><p className="font-bold text-sm">Produk</p><p className="text-xs opacity-70 hidden md:block">{inventory.length} SKU</p></div>
              </button>
              <button onClick={() => setAdminTab('TECHS')} className={`p-4 rounded-xl border flex flex-col md:flex-row items-center md:items-start gap-3 transition text-center md:text-left ${adminTab === 'TECHS' ? 'bg-slate-800 text-white shadow-lg' : 'bg-white hover:bg-slate-50'}`}>
                <div className={`p-2 rounded-full ${adminTab === 'TECHS' ? 'bg-slate-700' : 'bg-slate-100 text-slate-600'}`}><Users size={20} /></div>
                <div><p className="font-bold text-sm">Teknisi</p><p className="text-xs opacity-70 hidden md:block">{technicians.length} Staff</p></div>
              </button>
              
              {/* TOMBOL PENGATURAN BARU */}
              <button onClick={() => setShowSettingsModal(true)} className={`p-4 rounded-xl border flex flex-col md:flex-row items-center md:items-start gap-3 transition bg-blue-50 border-blue-200 hover:bg-blue-100 text-center md:text-left`}>
                <div className={`p-2 rounded-full bg-blue-200 text-blue-700`}><Settings size={20} /></div>
                <div><p className="font-bold text-sm text-blue-900">Pengaturan</p><p className="text-xs text-blue-700 hidden md:block">Database & Akun</p></div>
              </button>
            </div>

            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 md:p-6 overflow-hidden">
              {/* ADMIN: MANAJEMEN CABANG */}
              {adminTab === 'BRANCHES' && (
                <div>
                  <h2 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800"><Building size={24} /> Daftar Cabang Optik</h2>
                  <div className="bg-slate-50 p-4 rounded-lg border mb-6">
                    <h3 className="text-sm font-bold mb-3">Tambah Cabang Baru</h3>
                    <div className="flex flex-col md:flex-row gap-4 items-end">
                      <div className="flex-1 w-full"><label className="text-xs font-semibold block mb-1">Nama Cabang</label><input type="text" className="w-full p-2 text-sm border rounded" placeholder="Contoh: Cabang Fatmawati" value={newBranch.name} onChange={e => setNewBranch({...newBranch, name: e.target.value})} /></div>
                      <div className="w-full md:w-48"><label className="text-xs font-semibold block mb-1">Tipe Operasional</label><select className="w-full p-2 text-sm border rounded bg-white" value={newBranch.type} onChange={e => setNewBranch({...newBranch, type: e.target.value})}><option value="SALES">Sales Only</option><option value="PRODUCTION">Sales + Produksi</option></select></div>
                      <button onClick={handleAddBranch} className="w-full md:w-auto bg-slate-800 text-white px-4 py-2 text-sm rounded hover:bg-slate-900 font-bold flex items-center justify-center gap-2"><Plus size={16} /> Tambah</button>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-left border rounded overflow-hidden min-w-[600px]">
                      <thead className="bg-slate-100 font-bold"><tr><th className="p-3">ID</th><th className="p-3">Nama Cabang</th><th className="p-3">Tipe</th><th className="p-3 text-center">Status</th></tr></thead>
                      <tbody>
                        {branches.map(b => (
                          <tr key={b.id} className="border-t hover:bg-slate-50">
                            <td className="p-3 font-mono text-slate-500">{b.id}</td>
                            <td className="p-3 font-medium">{b.name}</td>
                            <td className="p-3"><span className={`text-xs px-2 py-1 rounded font-bold ${b.type === 'PRODUCTION' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'}`}>{b.type}</span></td>
                            <td className="p-3 text-center text-green-600 font-bold text-xs">AKTIF</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* ADMIN: MANAJEMEN INVENTORY */}
              {adminTab === 'INVENTORY' && (
                <div>
                   <h2 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800"><Package size={24} /> Master Data Produk (Admin)</h2>
                  <div className="bg-slate-50 p-4 rounded-lg border mb-6">
                    <h3 className="text-sm font-bold mb-3">Registrasi Item Baru (SKU)</h3>
                    <div className="grid grid-cols-1 md:grid-cols-5 gap-4 items-end">
                      <div className="md:col-span-2"><label className="text-xs font-semibold block mb-1">Nama Produk/Part</label><input type="text" className="w-full p-2 text-sm border rounded" placeholder="Nama Barang..." value={newItem.name} onChange={e => setNewItem({...newItem, name: e.target.value})} /></div>
                      <div><label className="text-xs font-semibold block mb-1">Kategori</label><select className="w-full p-2 text-sm border rounded bg-white" value={newItem.type} onChange={e => setNewItem({...newItem, type: e.target.value})}><option value="frame">Frame</option><option value="lens">Lensa</option><option value="part">Sparepart</option></select></div>
                       <div><label className="text-xs font-semibold block mb-1">Stok Awal (Pusat)</label><input type="number" className="w-full p-2 text-sm border rounded" placeholder="0" value={newItem.stock} onChange={e => setNewItem({...newItem, stock: parseInt(e.target.value) || 0})} /></div>
                      <div><label className="text-xs font-semibold block mb-1">Harga Satuan</label><input type="number" className="w-full p-2 text-sm border rounded" placeholder="Rp" value={newItem.price} onChange={e => setNewItem({...newItem, price: parseInt(e.target.value) || 0})} /></div>
                    </div>
                    <div className="mt-4 text-right"><button onClick={handleAddItem} className="w-full md:w-auto bg-slate-800 text-white px-6 py-2 text-sm rounded hover:bg-slate-900 font-bold inline-flex justify-center items-center gap-2"><Plus size={16} /> Simpan Master Data</button></div>
                  </div>
                  <div className="max-h-[500px] overflow-y-auto border rounded overflow-x-auto">
                     <table className="w-full text-sm text-left min-w-[600px]">
                      <thead className="bg-slate-100 font-bold sticky top-0"><tr><th className="p-3">ID</th><th className="p-3">Nama Item</th><th className="p-3">Kategori</th><th className="p-3">Stok Pusat</th><th className="p-3">Harga</th></tr></thead>
                      <tbody>
                        {inventory.map(item => (
                          <tr key={item.id} className="border-t hover:bg-slate-50">
                            <td className="p-3 font-mono text-slate-500 text-xs">{item.id}</td>
                            <td className="p-3 font-medium">{item.name}</td>
                            <td className="p-3 capitalize text-slate-500">{item.type}</td>
                            <td className="p-3 font-bold">{item.stock}</td>
                            <td className="p-3 text-slate-600">{formatRupiah(item.price)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* ADMIN: MANAJEMEN TEKNISI */}
              {adminTab === 'TECHS' && (
                <div>
                   <h2 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800"><Users size={24} /> Data Teknisi Produksi</h2>
                  <div className="bg-slate-50 p-4 rounded-lg border mb-6 flex flex-col md:flex-row gap-4 items-end">
                    <div className="flex-1 w-full"><label className="text-xs font-semibold block mb-1">Nama & Spesialisasi</label><input type="text" className="w-full p-2 text-sm border rounded" placeholder="Contoh: Rian (Finishing)" value={newTech} onChange={e => setNewTech(e.target.value)} /></div>
                    <button onClick={handleAddTech} className="w-full md:w-auto bg-slate-800 text-white px-4 py-2 text-sm rounded hover:bg-slate-900 font-bold flex items-center justify-center gap-2"><Plus size={16} /> Tambah Teknisi</button>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                    {technicians.map(t => (
                      <div key={t.id} className="p-4 border rounded-lg flex items-center gap-3 bg-white shadow-sm">
                        <div className="w-10 h-10 bg-slate-100 rounded-full flex items-center justify-center text-slate-500"><HardHat size={20} /></div>
                        <div><p className="font-bold text-sm text-slate-800">{t.name}</p><p className="text-xs text-slate-400">ID: {t.id}</p></div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
        
        {/* --- VIEW: POS (POINT OF SALE) --- */}
        {activeRole === 'POS' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-6">
              <div className="bg-white p-4 md:p-6 rounded-xl shadow-sm border border-slate-200">
                <div className="flex justify-between items-center mb-4"><h2 className="text-lg font-bold flex items-center gap-2 text-slate-700"><User size={20} /> Data Pelanggan & Resep</h2></div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                  <div><label className="block text-xs font-semibold text-slate-500 mb-1">Nama Pelanggan</label><input type="text" className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" placeholder="Contoh: Budi Santoso" value={customerName} onChange={(e) => setCustomerName(e.target.value)} /></div>
                  <div><label className="block text-xs font-semibold text-slate-500 mb-1">Kirim Order Ke (Produksi)</label><select className="w-full p-2 border rounded-lg bg-slate-50" value={selectedProductionHub} onChange={(e) => setSelectedProductionHub(e.target.value)}>{branches.filter(b => b.type === 'PRODUCTION').map(b => (<option key={b.id} value={b.id}>{b.name}</option>))}</select></div>
                </div>
                {/* Kartu Resep */}
                <div className="border rounded-lg overflow-hidden overflow-x-auto">
                  <table className="w-full text-sm text-center min-w-[300px]">
                    <thead className="bg-slate-100 text-slate-600 font-semibold"><tr><th className="p-2 text-left">Mata</th><th className="p-2">SPH</th><th className="p-2">CYL</th><th className="p-2">AXIS</th><th className="p-2">ADD</th></tr></thead>
                    <tbody className="divide-y">
                      <tr><td className="p-2 font-bold text-left bg-slate-50">R (OD)</td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="-0.00" value={prescription.od.sph} onChange={e => setPrescription({...prescription, od: {...prescription.od, sph: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="-0.00" value={prescription.od.cyl} onChange={e => setPrescription({...prescription, od: {...prescription.od, cyl: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="0" value={prescription.od.axis} onChange={e => setPrescription({...prescription, od: {...prescription.od, axis: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="+0.00" value={prescription.od.add} onChange={e => setPrescription({...prescription, od: {...prescription.od, add: e.target.value}})} /></td></tr>
                      <tr><td className="p-2 font-bold text-left bg-slate-50">L (OS)</td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="-0.00" value={prescription.os.sph} onChange={e => setPrescription({...prescription, os: {...prescription.os, sph: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="-0.00" value={prescription.os.cyl} onChange={e => setPrescription({...prescription, os: {...prescription.os, cyl: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="0" value={prescription.os.axis} onChange={e => setPrescription({...prescription, os: {...prescription.os, axis: e.target.value}})} /></td><td className="p-1"><input type="text" className="w-full text-center p-1 border rounded" placeholder="+0.00" value={prescription.os.add} onChange={e => setPrescription({...prescription, os: {...prescription.os, add: e.target.value}})} /></td></tr>
                    </tbody>
                  </table>
                  <div className="p-2 bg-slate-50 border-t flex items-center gap-2"><span className="text-sm font-semibold text-slate-600">PD (Pupil Distance):</span><input type="text" className="w-20 p-1 border rounded text-center" placeholder="62" value={prescription.pd} onChange={e => setPrescription({...prescription, pd: e.target.value})} /><span className="text-xs text-slate-500">mm</span></div>
                </div>
              </div>

              {/* Katalog Produk */}
              <div className="bg-white p-4 md:p-6 rounded-xl shadow-sm border border-slate-200">
                <h2 className="text-lg font-bold mb-4 flex items-center gap-2 text-slate-700"><Package size={20} /> Katalog Frame & Lensa</h2>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
                  {inventory.filter(i => ['frame', 'lens'].includes(i.type)).map(item => (
                    <button key={item.id} onClick={() => addToCart(item)} className="text-left border rounded-lg p-3 hover:border-teal-500 hover:shadow-md transition bg-slate-50">
                      <div className="text-sm font-bold text-slate-800 line-clamp-2 h-10">{item.name}</div>
                      <div className="text-xs text-slate-500 mb-2">Stok {currentBranch.name}: {getBranchQty(currentBranch.id, item.id)}</div>
                      <div className="text-teal-700 font-semibold">{formatRupiah(item.price)}</div>
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* Kolom Kanan: Keranjang */}
            <div className="lg:col-span-1">
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 sticky top-24">
                <div className="p-4 border-b bg-slate-50 rounded-t-xl"><h2 className="font-bold text-lg flex items-center gap-2"><ShoppingCart size={20} /> Ringkasan Order</h2></div>
                <div className="p-4 min-h-[300px] max-h-[500px] overflow-y-auto flex flex-col">
                  {cart.length === 0 ? (
                    <div className="flex-1 flex flex-col items-center justify-center text-slate-400 py-10"><ShoppingCart size={48} className="mb-2 opacity-20" /><p className="text-sm">Keranjang kosong</p></div>
                  ) : (
                    <div className="space-y-3 flex-1">{cart.map((item, idx) => (<div key={idx} className="flex justify-between items-center p-2 bg-slate-50 rounded border"><div><div className="text-sm font-medium">{item.name}</div><div className="text-xs text-teal-600">{formatRupiah(item.price)}</div></div><button onClick={() => removeFromCart(idx)} className="text-red-400 hover:text-red-600"><Plus size={16} className="rotate-45" /></button></div>))}</div>
                  )}
                  <div className="mt-4 pt-4 border-t">
                    <div className="flex justify-between font-bold text-lg mb-4"><span>Total</span><span>{formatRupiah(cart.reduce((a, b) => a + b.price, 0))}</span></div>
                    <button onClick={handleCheckout} className="w-full bg-teal-600 hover:bg-teal-700 text-white py-3 rounded-lg font-bold shadow-lg shadow-teal-200 transition flex items-center justify-center gap-2">Proses Pembayaran <ArrowRight size={18} /></button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* --- VIEW: PRODUCTION (LAB) --- */}
        {activeRole === 'PRODUCTION' && (
          <div className="space-y-6">
            <div className="flex flex-col md:flex-row items-center justify-between gap-4">
              <h2 className="text-2xl font-bold text-slate-800">Dashboard Produksi: {currentBranch.name}</h2>
              <div className="flex gap-2 w-full md:w-auto overflow-x-auto">
                <span className="px-3 py-1 bg-yellow-100 text-yellow-800 rounded-full text-xs font-bold whitespace-nowrap">Pending: {orders.filter(o => o.productionBranchId === currentBranch.id && o.status === 'PENDING').length}</span>
                <span className="px-3 py-1 bg-blue-100 text-blue-800 rounded-full text-xs font-bold whitespace-nowrap">In Progress: {orders.filter(o => o.productionBranchId === currentBranch.id && o.status === 'IN_PROGRESS').length}</span>
              </div>
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              {orders.filter(o => o.productionBranchId === currentBranch.id && o.status !== 'COMPLETED').map(order => (
                <div key={order.id} className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-4 border-b bg-slate-50 flex justify-between items-center">
                    <div><h3 className="font-bold text-lg">{order.customer}</h3><p className="text-xs text-slate-500">Ref: {order.id} • Dari: {branches.find(b => b.id === order.branchId)?.name}</p></div>
                    <div className={`px-3 py-1 rounded-full text-xs font-bold ${getStatusColor(order.status)}`}>{order.status}</div>
                  </div>
                  <div className="p-4 grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <div className="bg-blue-50 p-3 rounded-lg mb-4 text-xs"><p className="font-bold text-blue-800 mb-1">Resep Kacamata:</p><div className="grid grid-cols-2 gap-2"><div><span className="font-semibold">R:</span> {order.prescription.od.sph} / {order.prescription.od.cyl} x {order.prescription.od.axis}</div><div><span className="font-semibold">L:</span> {order.prescription.os.sph} / {order.prescription.os.cyl} x {order.prescription.os.axis}</div></div><p className="mt-1">PD: {order.prescription.pd}mm</p></div>
                      <div className="space-y-2 mb-4"><p className="text-xs font-bold text-slate-500 uppercase">Items:</p>{order.items.map((item, i) => (<div key={i} className="text-sm flex items-center gap-2"><CheckCircle size={14} className="text-teal-600" /> {item.name}</div>))}</div>
                      <div className="flex items-center gap-2 p-2 bg-slate-100 rounded text-sm"><HardHat size={16} className="text-slate-500" /><span className="font-semibold text-slate-600">Teknisi:</span><span className="font-bold text-slate-800">{order.technician || '- Belum Ditunjuk -'}</span></div>
                    </div>
                    <div className="border-l-0 md:border-l border-t md:border-t-0 pt-4 md:pt-0 pl-0 md:pl-4 border-slate-100">
                      <div className="mb-4">
                        <label className="text-xs font-bold text-slate-500 block mb-2">Update Status & Teknisi:</label>
                        <div className="space-y-2">
                          {order.status === 'PENDING' && (
                            <div className="flex gap-2">
                              <select id={`tech-select-${order.id}`} className="text-xs border p-2 rounded flex-1 bg-white"><option value="">-- Pilih Teknisi --</option>{technicians.map(t => (<option key={t.id} value={t.name}>{t.name}</option>))}</select>
                              <button onClick={() => {const select = document.getElementById(`tech-select-${order.id}`); assignTechnicianAndStart(order.id, select.value);}} className="btn-action bg-blue-600 text-white">Mulai</button>
                            </div>
                          )}
                          {order.status !== 'PENDING' && (
                            <div className="flex gap-2 flex-wrap">
                              {order.status === 'IN_PROGRESS' && (<button onClick={() => updateOrderStatus(order.id, 'QC')} className="btn-action bg-purple-600 text-white flex-1">Selesai & QC</button>)}
                              {order.status === 'QC' && (<button onClick={() => updateOrderStatus(order.id, 'READY')} className="btn-action bg-green-600 text-white flex-1">Lolos QC</button>)}
                              {order.status === 'READY' && (<button onClick={() => updateOrderStatus(order.id, 'COMPLETED')} className="btn-action bg-slate-800 text-white flex-1">Kirim ke Cabang Asal</button>)}
                            </div>
                          )}
                        </div>
                      </div>
                      <div className="mb-4">
                         <label className="text-xs font-bold text-slate-500 block mb-2">Pencatatan Part (Stok Cabang):</label>
                         <div className="flex gap-2 mb-2">
                            <select id={`part-select-${order.id}`} className="text-xs border p-2 rounded w-full bg-white">
                              {inventory.map(i => (
                                <option key={i.id} value={i.id}>{i.name} (Stok: {getBranchQty(currentBranch.id, i.id)})</option>
                              ))}
                            </select>
                         </div>
                         <div className="flex gap-2">
                           <button onClick={() => {const select = document.getElementById(`part-select-${order.id}`); useSparePart(order.id, select.value, false);}} className="text-xs bg-slate-200 hover:bg-slate-300 text-slate-800 px-2 py-2 rounded flex-1">+ Pakai Part</button>
                           <button onClick={() => {const select = document.getElementById(`part-select-${order.id}`); useSparePart(order.id, select.value, true);}} className="text-xs bg-red-100 hover:bg-red-200 text-red-700 px-2 py-2 rounded flex-1 flex items-center justify-center gap-1"><AlertTriangle size={12}/> Lapor Rusak</button>
                         </div>
                      </div>
                      <div className="bg-slate-50 rounded p-2 max-h-32 overflow-y-auto text-[10px] space-y-1 border">
                         {order.logs.map((log, idx) => (<div key={idx} className={`${log.type === 'danger' ? 'text-red-600 font-bold' : 'text-slate-600'}`}><span className="opacity-50">[{log.date.split(' ')[1]}]</span> {log.msg}</div>))}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
              {orders.filter(o => o.productionBranchId === currentBranch.id && o.status !== 'COMPLETED').length === 0 && (
                <div className="col-span-full py-12 text-center text-slate-400 bg-white rounded-xl border border-dashed border-slate-300"><ClipboardList size={48} className="mx-auto mb-2 opacity-20" /><p>Tidak ada pesanan aktif di antrian produksi.</p></div>
              )}
            </div>
          </div>
        )}

        {/* --- VIEW: WAREHOUSE (GUDANG & STOK CABANG) --- */}
        {activeRole === 'WAREHOUSE' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 bg-white rounded-xl shadow-sm border border-slate-200 p-4 md:p-6">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4">
                <div>
                  <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2"><Store size={24} /> Stok {currentBranch.name}</h2>
                  {currentBranch.id === 'GUDANG' && (<p className="text-sm text-slate-500 mt-1">Kelola fisik barang masuk dan keluar Gudang Pusat.</p>)}
                </div>
                {currentBranch.id !== 'GUDANG' && (<span className="text-xs bg-teal-50 text-teal-700 px-3 py-1 rounded-full border border-teal-200">Mode Cabang: Request ke Pusat</span>)}
                 {currentBranch.id === 'GUDANG' && (
                  <div className="flex gap-2 w-full md:w-auto">
                    <button onClick={() => setShowOpnameModal(true)} className="flex-1 md:flex-none flex items-center justify-center gap-2 px-3 py-2 bg-orange-100 text-orange-700 rounded-lg text-xs font-bold hover:bg-orange-200"><ClipboardCheck size={16} /> Opname Pusat</button>
                    <button onClick={() => setShowRestockModal(true)} className="flex-1 md:flex-none flex items-center justify-center gap-2 px-3 py-2 bg-teal-600 text-white rounded-lg text-xs font-bold hover:bg-teal-700 shadow-sm"><PackagePlus size={16} /> Restock Pusat</button>
                  </div>
                )}
              </div>
              
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left min-w-[500px]">
                  <thead className="bg-slate-50 text-slate-500 uppercase font-bold text-xs">
                    <tr><th className="p-3">Item</th><th className="p-3">Tipe</th><th className="p-3">Stok {currentBranch.id === 'GUDANG' ? 'Pusat' : 'Lokal'}</th>{currentBranch.id !== 'GUDANG' && <th className="p-3 text-right">Aksi</th>}</tr>
                  </thead>
                  <tbody className="divide-y">
                    {inventory.map(item => (
                      <tr key={item.id} className="hover:bg-slate-50">
                        <td className="p-3 font-medium text-slate-800">{item.name}</td>
                        <td className="p-3 text-slate-500 capitalize">{item.type}</td>
                        <td className="p-3">
                          {currentBranch.id === 'GUDANG' ? (
                            <span className={`font-bold ${item.stock < 10 ? 'text-red-500' : 'text-slate-700'}`}>{item.stock} pcs</span>
                          ) : (
                            <span className={`font-bold ${getBranchQty(currentBranch.id, item.id) < 5 ? 'text-red-500' : 'text-slate-700'}`}>{getBranchQty(currentBranch.id, item.id)} pcs</span>
                          )}
                        </td>
                        {currentBranch.id !== 'GUDANG' && (
                          <td className="p-3 text-right">
                            <button onClick={() => handleOpenRequestModal(item)} className="text-xs bg-teal-100 hover:bg-teal-200 text-teal-800 px-3 py-1 rounded transition whitespace-nowrap">Request ke Pusat</button>
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="lg:col-span-1">
              {currentBranch.id === 'GUDANG' ? (
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 md:p-6 sticky top-24">
                  <h2 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2"><Truck size={20} /> Permintaan Cabang</h2>
                  <div className="space-y-4">
                    {stockRequests.filter(r => r.status === 'PENDING').length === 0 ? (<p className="text-sm text-slate-400 text-center py-8">Tidak ada permintaan stok.</p>) : (
                      stockRequests.filter(r => r.status === 'PENDING').map(req => (
                        <div key={req.id} className="p-4 border rounded-lg bg-slate-50">
                          <div className="flex justify-between items-start mb-2"><div><p className="font-bold text-sm text-slate-800">{req.branchName}</p><p className="text-xs text-slate-500">Meminta: {req.itemName}</p></div><span className="font-bold text-teal-600 bg-teal-50 px-2 py-1 rounded text-xs">+{req.qty}</span></div>
                          {req.note && <p className="text-xs text-slate-400 italic mb-2">"{req.note}"</p>}
                          <div className="flex gap-2 mt-2"><button onClick={() => approveStockRequest(req.id)} className="flex-1 bg-teal-600 text-white text-xs py-2 rounded hover:bg-teal-700">Setujui & Kirim</button></div>
                        </div>
                      ))
                    )}
                  </div>
                  <div className="mt-8 pt-4 border-t">
                    <h3 className="text-sm font-bold text-slate-700 mb-2 flex items-center gap-2"><History size={16} /> Riwayat Pengiriman</h3>
                    <div className="space-y-2">{stockRequests.filter(r => r.status === 'APPROVED').map(req => (<div key={req.id} className="text-xs flex justify-between text-slate-500"><span>Kirim ke {req.branchName.substring(0, 15)}...</span><span className="text-green-600 font-bold">Selesai</span></div>))}</div>
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 md:p-6 sticky top-24">
                  <h2 className="text-lg font-bold text-slate-800 mb-4">Status Request Saya</h2>
                  <div className="space-y-3">
                    {stockRequests.filter(r => r.branchId === currentBranch.id).map(req => (
                      <div key={req.id} className="flex justify-between items-center text-sm border-b pb-2"><div><p className="font-medium">{req.itemName}</p><p className="text-xs text-slate-500">Qty: {req.qty}</p></div><span className={`text-xs px-2 py-1 rounded font-bold ${req.status === 'APPROVED' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'}`}>{req.status}</span></div>
                    ))}
                    {stockRequests.filter(r => r.branchId === currentBranch.id).length === 0 && (<p className="text-sm text-slate-400">Belum ada request aktif.</p>)}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
      <style>{`.btn-action {@apply px-3 py-1.5 rounded text-xs font-semibold shadow-sm hover:shadow transition;}`}</style>
    </div>
  );
}

<style>{`
  .btn-action {
    @apply px-3 py-1.5 rounded text-xs font-semibold shadow-sm hover:shadow transition;
  }
  .no-scrollbar::-webkit-scrollbar {
    display: none;
  }
  .no-scrollbar {
    -ms-overflow-style: none;
    scrollbar-width: none;
  }
  ::-webkit-scrollbar {
    width: 6px;
    height: 6px;
  }
  ::-webkit-scrollbar-track {
    background: #f1f5f9; 
    border-radius: 10px;
  }
  ::-webkit-scrollbar-thumb {
    background: #cbd5e1; 
    border-radius: 10px;
  }
  ::-webkit-scrollbar-thumb:hover {
    background: #94a3b8; 
  }
`}</style>