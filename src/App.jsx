import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  Wallet, TrendingDown, TrendingUp, Settings, LayoutDashboard, Plus, Receipt,
  CalendarDays, BarChart as BarChartIcon, Pencil, Check, X, Trash2, Download, Upload,
  Save, RefreshCw, ChevronLeft, ChevronRight, AlertCircle, CopyPlus, Info, Star, CreditCard, Zap, PiggyBank,
  PieChart as PieChartIcon, Search, Filter, Target, Database
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie } from 'recharts';

const CHART_COLORS = ['#3b82f6', '#ef4444', '#f59e0b', '#10b981', '#8b5cf6', '#ec4899', '#06b6d4', '#f97316'];

// --- 엑셀(XLSX) 다운로드 유틸리티 ---
const loadXlsxScript = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('Excel 라이브러리를 불러오는데 실패했습니다.'));
    document.head.appendChild(script);
  });
};

const exportToExcel = async (filename, rows) => {
  try {
    const XLSX = await loadXlsxScript();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
  } catch (error) {
    console.error("Excel Export Error:", error);
  }
};

// --- 공통 유틸리티 ---
const formatCurrency = (amount) => {
  if (amount === undefined || amount === null) return '0원';
  return new Intl.NumberFormat('ko-KR').format(amount) + '원';
};

// --- 로컬 데이터 로딩 ---
const loadLocalData = (key, fallback) => {
  try {
    const val = localStorage.getItem(key);
    return val ? JSON.parse(val) : fallback;
  } catch {
    return fallback;
  }
};

const saveLocalData = (key, data) => {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    console.error("Local storage save error", e);
  }
};

const defaultItems = [
  { id: 'i1', type: 'EXPENSE', category: '식비', name: '식대/간식', createdAt: new Date().toISOString() },
  { id: 'i2', type: 'EXPENSE', category: '교통비', name: '대중교통/주유', createdAt: new Date().toISOString() },
  { id: 'i3', type: 'EXPENSE', category: '주거/통신', name: '통신비/공과금', createdAt: new Date().toISOString() },
  { id: 'i4', type: 'INCOME', category: '주수입', name: '급여', createdAt: new Date().toISOString() },
  { id: 'i5', type: 'INCOME', category: '부수입', name: '보너스/용돈', createdAt: new Date().toISOString() },
  { id: 'i6', type: 'SAVING', category: '저축', name: '적금', createdAt: new Date().toISOString() },
  { id: 'i7', type: 'SAVING', category: '투자', name: '주식/코인', createdAt: new Date().toISOString() },
];

const defaultPaymentMethods = [
  { id: 'p1', name: '현금' },
  { id: 'p2', name: '신용카드' },
  { id: 'p3', name: '체크카드' },
  { id: 'p4', name: '계좌이체' },
  { id: 'p5', name: '자동이체' }
];

// --- [탭 1] 이달의 현황 대시보드 컴포넌트 ---
const DashboardTab = ({ items, paymentMethods, transactions, onImportFixedTransactions, budget, setBudget, onDeleteTransaction, goals }) => {
  const [currentMonthStr, setCurrentMonthStr] = useState(new Date().toISOString().slice(0, 7)); // YYYY-MM
  const [isEditingBudget, setIsEditingBudget] = useState(false);
  const [tempBudget, setTempBudget] = useState('');
  const [deletingTxId, setDeletingTxId] = useState(null);
  
  const changeMonth = (offset) => {
    const [y, m] = currentMonthStr.split('-');
    const d = new Date(parseInt(y, 10), parseInt(m, 10) - 1 + offset, 15);
    const newY = d.getFullYear();
    const newM = String(d.getMonth() + 1).padStart(2, '0');
    setCurrentMonthStr(`${newY}-${newM}`);
  };

  const monthlyTransactions = useMemo(() => {
    return transactions
      .filter(tx => tx.date.startsWith(currentMonthStr))
      .sort((a, b) => {
        if (a.date !== b.date) return b.date.localeCompare(a.date);
        return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
      });
  }, [transactions, currentMonthStr]);

  const stats = useMemo(() => {
    const income = monthlyTransactions.filter(t => t.type === 'INCOME').reduce((sum, t) => sum + t.amount, 0);
    const expense = monthlyTransactions.filter(t => t.type === 'EXPENSE').reduce((sum, t) => sum + t.amount, 0);
    const saving = monthlyTransactions.filter(t => t.type === 'SAVING').reduce((sum, t) => sum + t.amount, 0);
    return { income, expense, saving, balance: income - expense - saving };
  }, [monthlyTransactions]);

  const prevMonthStats = useMemo(() => {
    const [y, m] = currentMonthStr.split('-');
    const prevDate = new Date(parseInt(y, 10), parseInt(m, 10) - 2, 15);
    const prevMonthStr = prevDate.toISOString().slice(0, 7);
    
    const prevTxs = transactions.filter(tx => tx.date.startsWith(prevMonthStr));
    const income = prevTxs.filter(t => t.type === 'INCOME').reduce((sum, t) => sum + t.amount, 0);
    const expense = prevTxs.filter(t => t.type === 'EXPENSE').reduce((sum, t) => sum + t.amount, 0);
    return { income, expense };
  }, [transactions, currentMonthStr]);

  const renderTrend = (current, prev, type) => {
    if (prev === 0) return null;
    const diff = current - prev;
    const percent = Math.abs((diff / prev) * 100).toFixed(1);
    
    if (type === 'INCOME') {
      if (diff > 0) return <span className="text-[10px] text-blue-600 font-bold ml-1 bg-blue-50 px-1.5 py-0.5 rounded">▲ {percent}% 📈</span>;
      if (diff < 0) return <span className="text-[10px] text-red-600 font-bold ml-1 bg-red-50 px-1.5 py-0.5 rounded">▼ {percent}% 📉</span>;
    } else { 
      if (diff > 0) return <span className="text-[10px] text-red-600 font-bold ml-1 bg-red-50 px-1.5 py-0.5 rounded">▲ {percent}% 📈</span>;
      if (diff < 0) return <span className="text-[10px] text-blue-600 font-bold ml-1 bg-blue-50 px-1.5 py-0.5 rounded">▼ {percent}% 📉</span>;
    }
    return null;
  };

  const totalAccumulatedSavings = useMemo(() => {
    return transactions.filter(t => t.type === 'SAVING').reduce((sum, t) => sum + t.amount, 0);
  }, [transactions]);

  const missingFixedTransactions = useMemo(() => {
    const [year, month] = currentMonthStr.split('-');
    const prevMonthDate = new Date(parseInt(year, 10), parseInt(month, 10) - 2, 15);
    const prevMonthStr = prevMonthDate.toISOString().slice(0, 7);

    const prevFixedTxs = transactions.filter(tx => tx.date.startsWith(prevMonthStr) && tx.isFixed);
    return prevFixedTxs.filter(prevTx => {
      const originIdToMatch = prevTx.originalId || prevTx.id;
      const alreadyExists = monthlyTransactions.some(currTx => 
        (currTx.originalId === originIdToMatch || currTx.id === originIdToMatch)
      );
      return !alreadyExists;
    });
  }, [transactions, currentMonthStr, monthlyTransactions]);

  const handleBudgetSave = () => {
    const num = parseInt(tempBudget.replace(/[^0-9]/g, ''), 10);
    if (!isNaN(num)) {
      setBudget(num);
    }
    setIsEditingBudget(false);
  };

  const progress = budget > 0 ? Math.min((stats.expense / budget) * 100, 100) : 0;
  const isWarning = progress >= 80;

  return (
    <div className="space-y-6">
      <div className="bg-gradient-to-r from-emerald-500 to-teal-600 rounded-2xl p-6 shadow-lg text-white flex flex-col md:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-4">
          <div className="bg-white/20 p-4 rounded-full backdrop-blur-sm"><PiggyBank size={32} className="text-white" /></div>
          <div>
            <p className="text-emerald-100 font-medium text-sm">지금까지 모은 총 저축/투자 자산</p>
            <p className="text-3xl font-black tracking-tight">{formatCurrency(totalAccumulatedSavings)}</p>
          </div>
        </div>
        <div className="text-sm bg-white/10 px-4 py-2 rounded-lg backdrop-blur-sm">
          꾸준히 모인 금액은 계속 누적됩니다! ✨
        </div>
      </div>

      {goals.length > 0 && (
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <h3 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-2"><Target size={18} className="text-yellow-500" /> 나의 저축 목표 현황</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {goals.map(goal => {
              const saved = transactions.filter(t => t.goalId === goal.id).reduce((sum, t) => sum + t.amount, 0);
              const targetAmt = goal.targetAmount || 1; // 0 나누기 방지
              const progress = Math.min((saved / targetAmt) * 100, 100);
              return (
                <div key={goal.id} className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                  <div className="flex justify-between items-end mb-2">
                    <span className="font-bold text-slate-800 text-sm">{goal.name}</span>
                    <span className="font-bold text-slate-600 text-xs">{formatCurrency(saved)} / {formatCurrency(goal.targetAmount)}</span>
                  </div>
                  <div className="w-full bg-slate-200 rounded-full h-2.5 mb-1.5 overflow-hidden">
                    <div className="bg-yellow-400 h-2.5 rounded-full transition-all duration-500" style={{ width: `${progress}%` }}></div>
                  </div>
                  <div className="text-right text-[10px] font-bold text-slate-500">{progress.toFixed(1)}% 달성</div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {missingFixedTransactions.length > 0 && (
        <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 shadow-sm animate-in fade-in slide-in-from-top-4">
          <div className="flex items-center gap-3">
            <div className="bg-blue-100 text-blue-600 p-2 rounded-full shrink-0"><Info size={20} /></div>
            <div>
              <p className="text-sm font-bold text-blue-900">이전 달 고정 항목이 {missingFixedTransactions.length}건 있습니다!</p>
              <p className="text-xs text-blue-700 mt-0.5">매번 입력할 필요 없이 이번 달로 바로 복사해보세요.</p>
            </div>
          </div>
          <button onClick={() => onImportFixedTransactions(missingFixedTransactions, currentMonthStr)} className="w-full sm:w-auto px-4 py-2 bg-blue-600 text-white text-xs font-bold rounded-lg shadow-sm hover:bg-blue-700 transition flex items-center justify-center gap-2">
            <CopyPlus size={16} /> 지금 가져오기
          </button>
        </div>
      )}

      <div className="flex items-center justify-center gap-6 bg-white py-4 rounded-xl shadow-sm border border-slate-200">
        <button onClick={() => changeMonth(-1)} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><ChevronLeft size={24} className="text-slate-500"/></button>
        <span className="text-xl sm:text-2xl font-bold text-slate-800 min-w-[120px] text-center tracking-tight">
          {currentMonthStr.split('-')[0]}년 {currentMonthStr.split('-')[1]}월
        </span>
        <button onClick={() => changeMonth(1)} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><ChevronRight size={24} className="text-slate-500"/></button>
      </div>

      <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
        <div className="flex justify-between items-center mb-3">
          <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2"><Wallet size={16} className="text-slate-500"/> 이달의 지출 예산</h3>
          <div className="flex items-center gap-2">
            {isEditingBudget ? (
              <div className="flex items-center gap-1">
                <input type="text" value={tempBudget} onChange={(e) => setTempBudget(e.target.value.replace(/[^0-9]/g, ''))} autoFocus onBlur={handleBudgetSave} onKeyDown={(e) => e.key === 'Enter' && handleBudgetSave()} className="border-b-2 border-blue-500 outline-none w-24 text-right text-sm font-bold" placeholder="금액 입력" />
                <span className="text-sm font-bold text-slate-600">원</span>
              </div>
            ) : (
              <button onClick={() => { setTempBudget(budget.toString()); setIsEditingBudget(true); }} className="text-sm font-bold text-slate-600 hover:text-blue-600 flex items-center gap-1 group transition-colors">
                {formatCurrency(budget)} <Pencil size={12} className="text-slate-400 group-hover:text-blue-500"/>
              </button>
            )}
          </div>
        </div>
        <div className="relative w-full bg-slate-100 rounded-full h-4 overflow-hidden shadow-inner">
          <div className={`h-full rounded-full transition-all duration-500 ease-out ${isWarning ? 'bg-red-500' : 'bg-slate-800'}`} style={{ width: `${progress}%` }}></div>
        </div>
        <div className="flex justify-between mt-2">
          <span className="text-xs font-medium text-slate-500">순수 지출 {formatCurrency(stats.expense)} 사용됨</span>
          <span className={`text-xs font-bold ${isWarning ? 'text-red-500' : 'text-slate-800'}`}>{progress.toFixed(1)}%</span>
        </div>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex items-center gap-2 mb-2 text-blue-600">
            <TrendingUp size={16}/>
            <span className="text-xs font-bold text-slate-500">이번 달 수입</span>
          </div>
          <div className="flex items-center gap-2">
            <p className="text-lg font-bold text-blue-600 truncate">{formatCurrency(stats.income)}</p>
            {renderTrend(stats.income, prevMonthStats.income, 'INCOME')}
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex items-center gap-2 mb-2 text-red-600">
            <TrendingDown size={16}/>
            <span className="text-xs font-bold text-slate-500">이번 달 지출</span>
          </div>
          <div className="flex items-center gap-2">
            <p className="text-lg font-bold text-red-600 truncate">{formatCurrency(stats.expense)}</p>
            {renderTrend(stats.expense, prevMonthStats.expense, 'EXPENSE')}
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex items-center gap-2 mb-2 text-emerald-600"><PiggyBank size={16}/><span className="text-xs font-bold text-slate-500">이번 달 저축/투자</span></div>
          <p className="text-lg font-bold text-emerald-600 truncate">{formatCurrency(stats.saving)}</p>
        </div>
        <div className="bg-slate-800 p-4 rounded-xl shadow-sm text-white">
          <div className="flex items-center gap-2 mb-2"><Wallet size={16}/><span className="text-xs font-bold text-slate-300">이번 달 남은 돈</span></div>
          <p className="text-lg font-bold text-white truncate">{formatCurrency(stats.balance)}</p>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mt-6">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center">
          <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2"><Receipt size={18} className="text-slate-500" /> 이번 달 사용 내역</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 font-medium">일자</th>
                <th className="px-3 py-3 font-medium">유형</th>
                <th className="px-3 py-3 font-medium">항목명</th>
                <th className="px-3 py-3 font-medium">결제수단</th>
                <th className="px-3 py-3 font-medium text-right">금액</th>
                <th className="px-3 py-3 font-medium">메모/목표</th>
                <th className="px-3 py-3 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {monthlyTransactions.length === 0 ? (
                <tr><td colSpan="7" className="px-4 py-8 text-center text-slate-500">이번 달 기록된 내역이 없습니다.</td></tr>
              ) : (
                monthlyTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const pm = paymentMethods.find(p => p.id === tx.paymentMethodId);
                  const goal = goals.find(g => g.id === tx.goalId);
                  
                  let typeLabel = ''; let typeColor = ''; let amountColor = ''; let amountPrefix = '';
                  if (tx.type === 'INCOME') { typeLabel = '수입'; typeColor = 'bg-blue-100 text-blue-700'; amountColor = 'text-blue-600'; amountPrefix = '+'; } 
                  else if (tx.type === 'EXPENSE') { typeLabel = '지출'; typeColor = 'bg-red-100 text-red-700'; amountColor = 'text-red-600'; amountPrefix = '-'; } 
                  else { typeLabel = '저축'; typeColor = 'bg-emerald-100 text-emerald-700'; amountColor = 'text-emerald-600'; amountPrefix = '+'; }

                  return (
                    <tr key={tx.id} className="hover:bg-slate-50 transition-colors">
                      {deletingTxId === tx.id ? (
                        <>
                          <td colSpan="5" className="p-3 text-center text-red-600 font-bold text-sm bg-red-50/50">정말 삭제할까요?</td>
                          <td colSpan="2" className="p-3 text-center whitespace-nowrap bg-red-50/50">
                            <button onClick={() => { onDeleteTransaction(tx.id); setDeletingTxId(null); }} className="px-3 py-1.5 bg-red-600 text-white text-xs font-bold rounded-md mr-2 shadow-sm">삭제</button>
                            <button onClick={() => setDeletingTxId(null)} className="px-3 py-1.5 bg-slate-200 text-slate-700 text-xs font-bold rounded-md shadow-sm">취소</button>
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="px-3 py-3 text-slate-500 font-medium">{tx.date}</td>
                          <td className="px-3 py-3">
                            <span className={`px-2 py-1 rounded text-[10px] font-bold ${typeColor}`}>{typeLabel}</span>
                            {tx.isFixed && <span className="ml-1 text-[10px] text-slate-500 font-bold">📌고정</span>}
                          </td>
                          <td className="px-3 py-3 font-bold text-slate-800 truncate max-w-[120px]">{item ? `[${item.category}] ${item.name}` : '삭제된 항목'}</td>
                          <td className="px-3 py-3 text-slate-500">{pm ? pm.name : '-'}</td>
                          <td className={`px-3 py-3 text-right font-extrabold text-sm ${amountColor}`}>{amountPrefix}{formatCurrency(tx.amount)}</td>
                          <td className="px-3 py-3 text-slate-500 truncate max-w-[120px]">
                            {goal ? <span className="text-[10px] bg-yellow-100 text-yellow-800 font-bold px-1.5 py-0.5 rounded mr-1">🎯 {goal.name}</span> : null}
                            {tx.note || '-'}
                          </td>
                          <td className="px-3 py-3 text-center">
                            <button onClick={() => setDeletingTxId(tx.id)} className="p-1.5 text-slate-400 hover:text-red-600 bg-white rounded shadow-sm border border-slate-200 transition-colors" title="삭제"><Trash2 size={14}/></button>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [탭 2] 수입/지출 입력 컴포넌트 ---
const TransactionTab = ({ items, paymentMethods, transactions, quickAdds, setQuickAdds, goals, onAddTransaction, onDeleteTransaction }) => {
  const today = new Date().toISOString().split('T')[0];
  const defaultMethod = paymentMethods.length > 0 ? paymentMethods[0].id : '';
  const [formData, setFormData] = useState({ itemId: '', type: 'EXPENSE', amount: '', note: '', paymentMethodId: defaultMethod, date: today, isFixed: false, goalId: '' });
  const [deletingTxId, setDeletingTxId] = useState(null);
  const [inlineMessage, setInlineMessage] = useState(null);

  const filteredItems = useMemo(() => items.filter(item => item.type === formData.type).sort((a, b) => (a.category || '').localeCompare(b.category || '', 'ko-KR')), [items, formData.type]);
  const recentTransactions = useMemo(() => [...transactions].sort((a, b) => {
    if (a.date !== b.date) return b.date.localeCompare(a.date); 
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0); 
  }).slice(0, 30), [transactions]);

  const showMessage = (msg, type = 'success') => {
    setInlineMessage({ msg, type });
    setTimeout(() => setInlineMessage(null), 3000);
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!formData.itemId || !formData.amount || !formData.date) {
      showMessage('일자, 항목, 금액을 모두 입력해주세요.', 'error'); return;
    }
    const amt = parseInt(formData.amount.replace(/,/g, ''), 10);
    if (isNaN(amt) || amt <= 0) { showMessage('금액은 1원 이상이어야 합니다.', 'error'); return; }

    const newTx = {
      id: `t${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
      itemId: formData.itemId,
      type: formData.type,
      amount: amt,
      paymentMethodId: formData.paymentMethodId,
      date: formData.date, 
      note: formData.note,
      isFixed: formData.isFixed,
      goalId: formData.type === 'SAVING' ? formData.goalId : '', 
      createdAt: new Date().toISOString()
    };
    
    onAddTransaction(newTx);
    setFormData(prev => ({ ...prev, amount: '', note: '', isFixed: false, goalId: '' })); 
    showMessage('내역이 성공적으로 저장되었습니다!');
  };

  const handleAmountChange = (e) => {
    const rawValue = e.target.value.replace(/[^0-9]/g, '');
    if (!rawValue) { setFormData(prev => ({ ...prev, amount: '' })); return; }
    setFormData(prev => ({ ...prev, amount: new Intl.NumberFormat('ko-KR').format(parseInt(rawValue, 10)) }));
  };

  const saveAsQuickAdd = () => {
    if (!formData.itemId || !formData.amount) {
      showMessage('항목과 금액을 입력한 후 단축 버튼으로 만들 수 있습니다.', 'error'); return;
    }
    const item = items.find(i => i.id === formData.itemId);
    const amt = parseInt(formData.amount.replace(/,/g, ''), 10);
    const title = `${item ? item.name : '내역'} (${new Intl.NumberFormat('ko-KR').format(amt)})`;

    const newQa = {
      id: `q${Date.now()}`,
      title,
      type: formData.type,
      itemId: formData.itemId,
      amount: amt,
      paymentMethodId: formData.paymentMethodId,
      note: formData.note,
      goalId: formData.type === 'SAVING' ? formData.goalId : ''
    };
    setQuickAdds(prev => [...prev, newQa]);
    showMessage(`'${title}' 단축 버튼이 추가되었습니다!`);
  };

  const executeQuickAdd = (qa) => {
    const newTx = {
      id: `t${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
      itemId: qa.itemId,
      type: qa.type,
      amount: qa.amount,
      paymentMethodId: qa.paymentMethodId,
      date: new Date().toISOString().split('T')[0],
      note: qa.note,
      isFixed: false,
      goalId: qa.goalId || '',
      createdAt: new Date().toISOString()
    };
    onAddTransaction(newTx);
    showMessage(`${qa.title} 내역이 방금 등록되었습니다! (오늘 날짜)`);
  };

  const deleteQuickAdd = (id, e) => {
    e.stopPropagation();
    setQuickAdds(prev => prev.filter(q => q.id !== id));
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div className="lg:col-span-1 flex flex-col gap-4">
        
        {quickAdds.length > 0 && (
          <div className="bg-indigo-50 rounded-xl border border-indigo-100 p-4 shadow-sm">
            <h3 className="text-xs font-bold text-indigo-900 mb-2 flex items-center gap-1.5"><Zap size={14} className="fill-indigo-500 text-indigo-500"/> 1초 컷 자동 입력</h3>
            <div className="flex flex-wrap gap-2">
              {quickAdds.map(qa => (
                <div key={qa.id} className="relative group flex-shrink-0 max-w-full">
                  <button onClick={() => executeQuickAdd(qa)} className="pr-7 pl-3 py-1.5 bg-white border border-indigo-200 text-indigo-700 text-xs font-bold rounded-full shadow-sm hover:bg-indigo-100 hover:border-indigo-300 transition-all truncate max-w-full block">
                    {qa.title}
                  </button>
                  <button onClick={(e) => deleteQuickAdd(qa.id, e)} className="absolute right-1.5 top-1/2 -translate-y-1/2 p-0.5 text-slate-400 hover:text-red-500 bg-white rounded-full opacity-0 group-hover:opacity-100 transition-opacity">
                    <X size={12}/>
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4 sm:p-5 h-fit">
          <h2 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2"><Plus size={18} className="text-blue-600" /> 수입/지출 등록</h2>
          {inlineMessage && (
            <div className={`mb-4 p-3 text-xs rounded-lg flex items-center gap-2 ${inlineMessage.type === 'error' ? 'bg-red-50 text-red-600' : 'bg-emerald-50 text-emerald-700 font-bold'}`}>
              {inlineMessage.type === 'error' ? <AlertCircle size={14} className="shrink-0"/> : <Check size={14} className="shrink-0"/>}
              {inlineMessage.msg}
            </div>
          )}

          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">유형 선택</label>
              <div className="grid grid-cols-3 gap-2">
                <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'INCOME', itemId: '', goalId: '' }))} className={`py-2 px-1 rounded-lg border text-[13px] font-bold transition-all flex items-center justify-center gap-1 ${formData.type === 'INCOME' ? 'bg-blue-600 border-blue-600 text-white shadow-md shadow-blue-200' : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}`}>
                  {formData.type === 'INCOME' && <Check size={12} className="shrink-0"/>} 수입
                </button>
                <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'EXPENSE', itemId: '', goalId: '' }))} className={`py-2 px-1 rounded-lg border text-[13px] font-bold transition-all flex items-center justify-center gap-1 ${formData.type === 'EXPENSE' ? 'bg-red-600 border-red-600 text-white shadow-md shadow-red-200' : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}`}>
                  {formData.type === 'EXPENSE' && <Check size={12} className="shrink-0"/>} 지출
                </button>
                <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'SAVING', itemId: '', goalId: '' }))} className={`py-2 px-1 rounded-lg border text-[13px] font-bold transition-all flex items-center justify-center gap-1 ${formData.type === 'SAVING' ? 'bg-emerald-500 border-emerald-500 text-white shadow-md shadow-emerald-200' : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}`}>
                  {formData.type === 'SAVING' && <Check size={12} className="shrink-0"/>} 저축/투자
                </button>
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">일자</label>
              <input type="date" required value={formData.date} onChange={(e) => setFormData(prev => ({ ...prev, date: e.target.value }))} className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200 bg-white" />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">항목 (카테고리)</label>
              <select value={formData.itemId} onChange={(e) => setFormData(prev => ({ ...prev, itemId: e.target.value }))} className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200 bg-white font-medium">
                <option value="">-- 항목을 선택하세요 --</option>
                {filteredItems.map(item => <option key={item.id} value={item.id}>[{item.category}] {item.name}</option>)}
              </select>
              {filteredItems.length === 0 && <p className="text-[10px] text-red-500 mt-1">이 유형의 항목이 없습니다. [항목 관리]에서 먼저 추가해주세요.</p>}
            </div>

            {formData.type === 'SAVING' && goals.length > 0 && (
              <div className="bg-yellow-50 p-3 rounded-lg border border-yellow-200 animate-in fade-in slide-in-from-top-2">
                <label className="block text-xs font-bold text-yellow-800 mb-1 flex items-center gap-1"><Target size={12}/> 저축 목표 할당 (선택)</label>
                <select value={formData.goalId || ''} onChange={(e) => setFormData(prev => ({ ...prev, goalId: e.target.value }))} className="w-full rounded-lg border-yellow-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-yellow-400 bg-white">
                  <option value="">-- 목표 선택안함 --</option>
                  {goals.map(goal => <option key={goal.id} value={goal.id}>[목표] {goal.name} (목표액: {formatCurrency(goal.targetAmount)})</option>)}
                </select>
              </div>
            )}

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1 flex items-center gap-1"><CreditCard size={12}/> {formData.type === 'SAVING' ? '출금 수단' : '결제 수단'}</label>
              <select value={formData.paymentMethodId} onChange={(e) => setFormData(prev => ({ ...prev, paymentMethodId: e.target.value }))} className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200 bg-white">
                {paymentMethods.length === 0 && <option value="">수단 없음</option>}
                {paymentMethods.map(pm => <option key={pm.id} value={pm.id}>{pm.name}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">금액</label>
              <div className="relative">
                <input type="text" inputMode="numeric" required value={formData.amount} onChange={handleAmountChange} placeholder="0" className="w-full rounded-lg border-slate-200 border px-3 py-2 text-base font-bold focus:outline-none focus:ring-2 focus:ring-blue-200 text-right pr-8" />
                <span className="absolute right-3 top-2.5 text-sm font-bold text-slate-400">원</span>
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">메모 (선택)</label>
              <input type="text" value={formData.note} onChange={(e) => setFormData(prev => ({ ...prev, note: e.target.value }))} placeholder="상세 내용 등 간단한 메모" className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-200" />
            </div>
            <div className="flex items-center gap-2 mt-2 p-3 bg-slate-50 border border-slate-100 rounded-lg">
               <input type="checkbox" id="isFixed" checked={formData.isFixed} onChange={(e) => setFormData(prev => ({ ...prev, isFixed: e.target.checked }))} className="rounded text-blue-600 w-4 h-4" />
               <label htmlFor="isFixed" className="text-xs font-bold text-slate-700 cursor-pointer select-none">매월 고정 항목으로 설정 (자동 복사 지원)</label>
            </div>

            <div className="pt-2 flex flex-col gap-2">
              <button type="submit" className="w-full text-white text-sm font-bold py-3 rounded-lg bg-slate-800 hover:bg-slate-700 transition-colors shadow-md">
                입력하기
              </button>
              <button type="button" onClick={saveAsQuickAdd} className="w-full text-indigo-700 border border-indigo-200 text-xs font-bold py-2.5 rounded-lg bg-white hover:bg-indigo-50 transition-colors flex items-center justify-center gap-1.5 shadow-sm">
                <Star size={14} className="fill-indigo-400 text-indigo-400"/> 현재 내용을 단축 버튼으로 저장
              </button>
            </div>
          </form>
        </div>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden h-fit">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center">
          <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2"><Receipt size={18} className="text-slate-500" /> 최근 입력 내역</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 font-medium">일자</th>
                <th className="px-3 py-3 font-medium">유형</th>
                <th className="px-3 py-3 font-medium">항목명</th>
                <th className="px-3 py-3 font-medium">결제수단</th>
                <th className="px-3 py-3 font-medium text-right">금액</th>
                <th className="px-3 py-3 font-medium">메모/목표</th>
                <th className="px-3 py-3 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {recentTransactions.length === 0 ? (
                <tr><td colSpan="7" className="px-4 py-8 text-center text-slate-500">기록된 내역이 없습니다.</td></tr>
              ) : (
                recentTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const pm = paymentMethods.find(p => p.id === tx.paymentMethodId);
                  const goal = goals.find(g => g.id === tx.goalId);
                  
                  let typeLabel = ''; let typeColor = ''; let amountColor = ''; let amountPrefix = '';
                  if (tx.type === 'INCOME') { typeLabel = '수입'; typeColor = 'bg-blue-100 text-blue-700'; amountColor = 'text-blue-600'; amountPrefix = '+'; } 
                  else if (tx.type === 'EXPENSE') { typeLabel = '지출'; typeColor = 'bg-red-100 text-red-700'; amountColor = 'text-red-600'; amountPrefix = '-'; } 
                  else { typeLabel = '저축'; typeColor = 'bg-emerald-100 text-emerald-700'; amountColor = 'text-emerald-600'; amountPrefix = '+'; }

                  return (
                    <tr key={tx.id} className="hover:bg-slate-50 transition-colors">
                      {deletingTxId === tx.id ? (
                        <>
                          <td colSpan="5" className="p-3 text-center text-red-600 font-bold text-sm bg-red-50/50">정말 삭제할까요?</td>
                          <td colSpan="2" className="p-3 text-center whitespace-nowrap bg-red-50/50">
                            <button onClick={() => { onDeleteTransaction(tx.id); setDeletingTxId(null); }} className="px-3 py-1.5 bg-red-600 text-white text-xs font-bold rounded-md mr-2 shadow-sm">삭제</button>
                            <button onClick={() => setDeletingTxId(null)} className="px-3 py-1.5 bg-slate-200 text-slate-700 text-xs font-bold rounded-md shadow-sm">취소</button>
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="px-3 py-3 text-slate-500 font-medium">{tx.date}</td>
                          <td className="px-3 py-3">
                            <span className={`px-2 py-1 rounded text-[10px] font-bold ${typeColor}`}>{typeLabel}</span>
                            {tx.isFixed && <span className="ml-1 text-[10px] text-slate-500 font-bold">📌고정</span>}
                          </td>
                          <td className="px-3 py-3 font-bold text-slate-800 truncate max-w-[120px]">{item ? `[${item.category}] ${item.name}` : '삭제된 항목'}</td>
                          <td className="px-3 py-3 text-slate-500">{pm ? pm.name : '-'}</td>
                          <td className={`px-3 py-3 text-right font-extrabold text-sm ${amountColor}`}>
                            {amountPrefix}{formatCurrency(tx.amount)}
                          </td>
                          <td className="px-3 py-3 text-slate-500 truncate max-w-[120px]">
                            {goal ? <span className="text-[10px] bg-yellow-100 text-yellow-800 font-bold px-1.5 py-0.5 rounded mr-1">🎯 {goal.name}</span> : null}
                            {tx.note || '-'}
                          </td>
                          <td className="px-3 py-3 text-center">
                            <button onClick={() => setDeletingTxId(tx.id)} className="p-1.5 text-slate-400 hover:text-red-600 bg-white rounded shadow-sm border border-slate-200 transition-colors" title="삭제"><Trash2 size={14}/></button>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [탭 3] 통계/상세조회 컴포넌트 ---
const ReportTab = ({ items, paymentMethods, transactions, onDeleteTransaction, goals }) => {
  const [viewMode, setViewMode] = useState('monthly');
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedMonth, setSelectedMonth] = useState(new Date().toISOString().slice(0, 7));
  const [deletingTxId, setDeletingTxId] = useState(null);
  
  const [isFilterOpen, setIsFilterOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');

  const filteredTransactions = useMemo(() => {
    return transactions.filter(tx => {
      if (viewMode === 'daily' && tx.date !== selectedDate) return false;
      if (viewMode === 'monthly' && !tx.date.startsWith(selectedMonth)) return false;
      
      if (searchTerm) {
        const item = items.find(i => i.id === tx.itemId);
        const searchTarget = `${item?.category || ''} ${item?.name || ''} ${tx.note || ''}`.toLowerCase();
        if (!searchTarget.includes(searchTerm.toLowerCase())) return false;
      }

      if (minAmount && tx.amount < parseInt(minAmount, 10)) return false;
      if (maxAmount && tx.amount > parseInt(maxAmount, 10)) return false;

      return true;
    }).sort((a, b) => {
      if (a.date !== b.date) return b.date.localeCompare(a.date);
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });
  }, [transactions, viewMode, selectedDate, selectedMonth, searchTerm, minAmount, maxAmount, items]);

  const periodStats = useMemo(() => {
    return {
      income: filteredTransactions.filter(t => t.type === 'INCOME').reduce((sum, t) => sum + t.amount, 0),
      expense: filteredTransactions.filter(t => t.type === 'EXPENSE').reduce((sum, t) => sum + t.amount, 0),
      saving: filteredTransactions.filter(t => t.type === 'SAVING').reduce((sum, t) => sum + t.amount, 0)
    };
  }, [filteredTransactions]);

  const expenseByCategory = useMemo(() => {
    const expenses = filteredTransactions.filter(t => t.type === 'EXPENSE');
    const total = periodStats.expense;
    const grouped = {};
    expenses.forEach(tx => {
      const item = items.find(i => i.id === tx.itemId);
      const cat = item ? item.category : '미분류';
      grouped[cat] = (grouped[cat] || 0) + tx.amount;
    });
    return Object.entries(grouped)
      .map(([category, amount]) => ({ category, amount, percentage: total > 0 ? (amount / total) * 100 : 0 }))
      .sort((a, b) => b.amount - a.amount);
  }, [filteredTransactions, items, periodStats.expense]);

  return (
    <div className="space-y-6">
      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
        <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2"><BarChartIcon size={20} className="text-blue-600" /> 상세 내역 조회</h2>
        
        <div className="flex flex-col sm:flex-row items-center gap-4 w-full md:w-auto">
          <button onClick={() => setIsFilterOpen(!isFilterOpen)} className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors flex items-center gap-1.5 border ${isFilterOpen ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
            <Filter size={14} /> 고급 검색/필터
          </button>
          <div className="flex bg-slate-100 p-1 rounded-lg w-full sm:w-auto shadow-inner">
            <button onClick={() => setViewMode('daily')} className={`flex-1 sm:flex-none px-4 py-1.5 text-xs font-bold rounded-md transition-all ${viewMode === 'daily' ? 'bg-white shadow-sm text-blue-700' : 'text-slate-500 hover:text-slate-700'}`}>일별 조회</button>
            <button onClick={() => setViewMode('monthly')} className={`flex-1 sm:flex-none px-4 py-1.5 text-xs font-bold rounded-md transition-all ${viewMode === 'monthly' ? 'bg-white shadow-sm text-blue-700' : 'text-slate-500 hover:text-slate-700'}`}>월별 조회</button>
          </div>
          <div className="flex items-center gap-2 w-full sm:w-auto">
            <CalendarDays size={16} className="text-slate-400 hidden sm:block" />
            {viewMode === 'daily' ? (
              <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} className="w-full rounded-lg border-slate-200 border px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-200" />
            ) : (
              <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="w-full rounded-lg border-slate-200 border px-3 py-1.5 text-sm focus:ring-2 focus:ring-blue-200" />
            )}
          </div>
        </div>
      </div>

      {isFilterOpen && (
        <div className="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100 shadow-sm grid grid-cols-1 sm:grid-cols-3 gap-4 animate-in fade-in slide-in-from-top-2">
          <div>
            <label className="block text-xs font-bold text-indigo-900 mb-1">단어 검색</label>
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-2.5 text-indigo-400" />
              <input type="text" value={searchTerm} onChange={e => setSearchTerm(e.target.value)} placeholder="내역, 항목, 메모 등 입력" className="w-full pl-8 pr-3 py-2 text-sm border border-indigo-200 rounded-lg focus:ring-2 focus:ring-indigo-300 outline-none bg-white" />
            </div>
          </div>
          <div>
            <label className="block text-xs font-bold text-indigo-900 mb-1">최소 금액</label>
            <input type="number" value={minAmount} onChange={e => setMinAmount(e.target.value)} placeholder="예: 10000" className="w-full px-3 py-2 text-sm border border-indigo-200 rounded-lg focus:ring-2 focus:ring-indigo-300 outline-none bg-white" />
          </div>
          <div>
            <label className="block text-xs font-bold text-indigo-900 mb-1">최대 금액</label>
            <input type="number" value={maxAmount} onChange={e => setMaxAmount(e.target.value)} placeholder="예: 50000" className="w-full px-3 py-2 text-sm border border-indigo-200 rounded-lg focus:ring-2 focus:ring-indigo-300 outline-none bg-white" />
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between gap-4">
          <div><p className="text-xs text-slate-500 font-bold mb-1">조회 기간 총 수입</p><p className="text-xl font-black text-blue-600">{formatCurrency(periodStats.income)}</p></div>
          <div className="p-3 bg-blue-50 text-blue-500 rounded-full shrink-0"><TrendingUp size={24} /></div>
        </div>
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between gap-4">
          <div><p className="text-xs text-slate-500 font-bold mb-1">조회 기간 총 지출</p><p className="text-xl font-black text-red-600">{formatCurrency(periodStats.expense)}</p></div>
          <div className="p-3 bg-red-50 text-red-500 rounded-full shrink-0"><TrendingDown size={24} /></div>
        </div>
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between gap-4">
          <div><p className="text-xs text-slate-500 font-bold mb-1">조회 기간 저축/투자</p><p className="text-xl font-black text-emerald-600">{formatCurrency(periodStats.saving)}</p></div>
          <div className="p-3 bg-emerald-50 text-emerald-500 rounded-full shrink-0"><PiggyBank size={24} /></div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50">
          <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
            <PieChartIcon size={18} className="text-slate-500" /> 지출 카테고리 비중
          </h2>
        </div>
        <div className="p-5">
          {expenseByCategory.length > 0 ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {expenseByCategory.map((item, idx) => (
                <div key={idx} className="flex justify-between items-center p-3.5 bg-slate-50 rounded-xl border border-slate-100 hover:border-slate-200 transition-colors">
                  <div className="flex items-center gap-2.5">
                    <span className="w-3 h-3 rounded-full shadow-sm" style={{ backgroundColor: CHART_COLORS[idx % CHART_COLORS.length] }}></span>
                    <span className="text-sm font-bold text-slate-700">{item.category}</span>
                  </div>
                  <div className="text-right flex flex-col justify-center">
                    <span className="text-sm font-extrabold text-slate-800">{formatCurrency(item.amount)}</span>
                    <span className="text-[11px] font-bold text-slate-400">{item.percentage.toFixed(1)}%</span>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="text-center text-slate-400 text-xs py-8">지출 데이터가 없습니다.</div>
          )}
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center">
          <h2 className="text-sm font-bold text-slate-800">
            {viewMode === 'daily' ? `${selectedDate} 상세 내역` : `${selectedMonth.replace('-', '년 ')}월 상세 내역`}
            {(searchTerm || minAmount || maxAmount) && <span className="ml-2 text-[10px] bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full">검색 필터 적용됨</span>}
          </h2>
          <button onClick={() => {
              const rows = [['일자', '유형', '항목', '결제수단', '금액', '할당된 목표', '메모'], ...filteredTransactions.map(tx => {
                const item = items.find(i => i.id === tx.itemId);
                const pm = paymentMethods.find(p => p.id === tx.paymentMethodId);
                const goal = goals.find(g => g.id === tx.goalId);
                const typeStr = tx.type === 'INCOME' ? '수입' : tx.type === 'EXPENSE' ? '지출' : '저축';
                return [tx.date, typeStr, item ? `[${item.category}] ${item.name}` : '-', pm ? pm.name : '-', tx.amount, goal ? goal.name : '-', tx.note || ''];
              })];
              exportToExcel(`가계부_상세내역`, rows);
            }} 
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors"
          >
            <Download size={14} /> 엑셀 다운로드
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 font-medium text-center w-12">순번</th>
                {viewMode === 'monthly' && <th className="px-3 py-3 font-bold text-blue-700 bg-blue-50/50">일자(MM-DD)</th>}
                <th className="px-3 py-3 font-medium">유형</th>
                <th className="px-3 py-3 font-medium">항목명</th>
                <th className="px-3 py-3 font-medium">결제수단</th>
                <th className="px-3 py-3 font-medium text-right">금액</th>
                <th className="px-3 py-3 font-medium">메모/목표</th>
                <th className="px-3 py-3 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredTransactions.length === 0 ? (
                <tr><td colSpan={viewMode === 'monthly' ? 8 : 7} className="px-4 py-8 text-center text-slate-500">해당 조건의 내역이 없습니다.</td></tr>
              ) : (
                filteredTransactions.map((tx, index) => {
                  const item = items.find(i => i.id === tx.itemId);
                  const pm = paymentMethods.find(p => p.id === tx.paymentMethodId);
                  const goal = goals.find(g => g.id === tx.goalId);
                  const seqNum = filteredTransactions.length - index; 

                  let typeLabel = ''; let typeColor = ''; let amountColor = ''; let amountPrefix = '';
                  if (tx.type === 'INCOME') { typeLabel = '수입'; typeColor = 'bg-blue-100 text-blue-700'; amountColor = 'text-blue-600'; amountPrefix = '+'; } 
                  else if (tx.type === 'EXPENSE') { typeLabel = '지출'; typeColor = 'bg-red-100 text-red-700'; amountColor = 'text-red-600'; amountPrefix = '-'; } 
                  else { typeLabel = '저축'; typeColor = 'bg-emerald-100 text-emerald-700'; amountColor = 'text-emerald-600'; amountPrefix = '+'; }

                  return (
                    <tr key={tx.id} className="hover:bg-slate-50 transition-colors">
                      {deletingTxId === tx.id ? (
                        <>
                          <td colSpan={viewMode === 'monthly' ? 6 : 5} className="p-3 text-center text-red-600 font-bold text-sm bg-red-50/50">정말 삭제하시겠습니까?</td>
                          <td colSpan="2" className="p-3 text-center whitespace-nowrap bg-red-50/50">
                            <button onClick={() => { onDeleteTransaction(tx.id); setDeletingTxId(null); }} className="px-3 py-1.5 bg-red-600 text-white text-xs font-bold rounded-md mr-2 shadow-sm">삭제</button>
                            <button onClick={() => setDeletingTxId(null)} className="px-3 py-1.5 bg-slate-200 text-slate-700 text-xs font-bold rounded-md shadow-sm">취소</button>
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="px-3 py-3 text-slate-400 text-center">{seqNum}</td>
                          {viewMode === 'monthly' && <td className="px-3 py-3 font-black text-slate-800 bg-slate-50/30">{tx.date.substring(5)}</td>}
                          <td className="px-3 py-3">
                            <span className={`px-2 py-1 rounded text-[10px] font-bold ${typeColor}`}>
                              {typeLabel}
                            </span>
                          </td>
                          <td className="px-3 py-3 font-bold text-slate-800 truncate max-w-[120px]">{item ? `[${item.category}] ${item.name}` : '삭제된 항목'}</td>
                          <td className="px-3 py-3 text-slate-500">{pm ? pm.name : '-'}</td>
                          <td className={`px-3 py-3 text-right font-extrabold text-sm ${amountColor}`}>
                            {amountPrefix}{formatCurrency(tx.amount)}
                          </td>
                          <td className="px-3 py-3 text-slate-500 truncate max-w-[150px]">
                            {goal ? <span className="text-[10px] bg-yellow-100 text-yellow-800 font-bold px-1.5 py-0.5 rounded mr-1">🎯 {goal.name}</span> : null}
                            {tx.note || '-'}
                          </td>
                          <td className="px-3 py-3 text-center">
                            <button onClick={() => setDeletingTxId(tx.id)} className="p-1.5 text-slate-400 hover:text-red-600 bg-white rounded shadow-sm border border-slate-200 transition-colors" title="삭제"><Trash2 size={14}/></button>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [탭 4] 항목 & 결제수단 관리 컴포넌트 ---
const ItemManagementTab = ({ items, paymentMethods, goals, onAddItem, onUpdateItem, onDeleteItem, onAddPM, onDeletePM, onAddGoal, onDeleteGoal }) => {
  const [newItem, setNewItem] = useState({ type: 'EXPENSE', category: '', name: '' });
  const [newPM, setNewPM] = useState('');
  const [newGoal, setNewGoal] = useState({ name: '', targetAmount: '' });
  
  const [editingItemId, setEditingItemId] = useState(null);
  const [editFormData, setEditFormData] = useState({});
  const [deletingItemId, setDeletingItemId] = useState(null);
  const [inlineMessage, setInlineMessage] = useState(null);

  const showMessage = (msg, type = 'success') => {
    setInlineMessage({ msg, type }); setTimeout(() => setInlineMessage(null), 4000);
  };

  const sortedItems = useMemo(() => [...items].sort((a, b) => {
    const typeOrder = { 'INCOME': 1, 'EXPENSE': 2, 'SAVING': 3 };
    if (a.type !== b.type) return typeOrder[a.type] - typeOrder[b.type];
    return (a.category || '').localeCompare(b.category || '', 'ko-KR');
  }), [items]);

  const handleAdd = (e) => {
    e.preventDefault();
    if (!newItem.category || !newItem.name) return;
    onAddItem({ id: `i${Date.now()}_${Math.random().toString(36).substr(2, 5)}`, ...newItem, createdAt: new Date().toISOString() });
    setNewItem({ type: 'EXPENSE', category: '', name: '' });
    showMessage('새 항목이 추가되었습니다.');
  };

  const handleAddPMSubmit = (e) => {
    e.preventDefault();
    if (!newPM.trim()) return;
    onAddPM({ id: `p${Date.now()}`, name: newPM.trim() });
    setNewPM('');
    showMessage('결제 수단이 추가되었습니다.');
  };

  const handleAddGoalSubmit = (e) => {
    e.preventDefault();
    if (!newGoal.name.trim() || !newGoal.targetAmount) return;
    const targetAmt = parseInt(newGoal.targetAmount.replace(/,/g, ''), 10);
    if (isNaN(targetAmt) || targetAmt <= 0) { showMessage('목표 금액을 올바르게 입력해주세요.', 'error'); return; }
    
    onAddGoal({ id: `g${Date.now()}`, name: newGoal.name.trim(), targetAmount: targetAmt, createdAt: new Date().toISOString() });
    setNewGoal({ name: '', targetAmount: '' });
    showMessage('새로운 저축 목표가 생성되었습니다!');
  };

  const handleSaveEdit = (id) => {
    if (!editFormData.name) return;
    onUpdateItem({ ...items.find(i => i.id === id), ...editFormData, updatedAt: new Date().toISOString() });
    setEditingItemId(null); 
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div className="lg:col-span-1 flex flex-col gap-6">
        {inlineMessage && (
          <div className={`p-3 text-xs font-bold rounded-lg flex items-center gap-2 ${inlineMessage.type === 'error' ? 'bg-red-50 text-red-600 border border-red-200' : 'bg-blue-50 text-blue-700 border border-blue-200'}`}>
            {inlineMessage.type === 'error' ? <AlertCircle size={16} className="shrink-0"/> : <Check size={16} className="shrink-0"/>}
            {inlineMessage.msg}
          </div>
        )}

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4 sm:p-5">
          <h2 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2"><Plus size={18} className="text-blue-600" /> 새 항목 등록</h2>
          <form onSubmit={handleAdd} className="space-y-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">유형</label>
              <div className="grid grid-cols-3 bg-slate-100 p-1 rounded-lg border border-slate-200">
                <button type="button" onClick={() => setNewItem({...newItem, type: 'INCOME'})} className={`py-1.5 text-xs font-bold rounded-md transition-all ${newItem.type === 'INCOME' ? 'bg-white shadow-sm text-blue-600' : 'text-slate-500'}`}>수입</button>
                <button type="button" onClick={() => setNewItem({...newItem, type: 'EXPENSE'})} className={`py-1.5 text-xs font-bold rounded-md transition-all ${newItem.type === 'EXPENSE' ? 'bg-white shadow-sm text-red-600' : 'text-slate-500'}`}>지출</button>
                <button type="button" onClick={() => setNewItem({...newItem, type: 'SAVING'})} className={`py-1.5 text-xs font-bold rounded-md transition-all ${newItem.type === 'SAVING' ? 'bg-white shadow-sm text-emerald-600' : 'text-slate-500'}`}>저축</button>
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">카테고리 (대분류)</label>
              <input type="text" required value={newItem.category} onChange={e => setNewItem({...newItem, category: e.target.value})} placeholder="예: 식비, 통신비, 급여, 저축" className="w-full border border-slate-200 p-2 rounded-lg text-sm focus:ring-2 focus:ring-blue-200 outline-none" />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">세부 항목명</label>
              <input type="text" required value={newItem.name} onChange={e => setNewItem({...newItem, name: e.target.value})} placeholder="예: 외식, 배달, 넷플릭스, 청약" className="w-full border border-slate-200 p-2 rounded-lg text-sm focus:ring-2 focus:ring-blue-200 outline-none" />
            </div>
            <button type="submit" className="w-full bg-slate-800 text-white py-2.5 rounded-lg hover:bg-slate-700 text-xs font-bold transition-colors shadow-sm">항목 추가하기</button>
          </form>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4 sm:p-5">
          <h2 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2"><CreditCard size={18} className="text-blue-600" /> 결제 수단 관리</h2>
          <form onSubmit={handleAddPMSubmit} className="flex gap-2 mb-4">
            <input type="text" required value={newPM} onChange={e => setNewPM(e.target.value)} placeholder="새 결제수단 (예: A카드)" className="flex-1 border border-slate-200 p-2 rounded-lg text-sm focus:ring-2 focus:ring-blue-200 outline-none" />
            <button type="submit" className="bg-slate-800 text-white px-4 py-2 rounded-lg hover:bg-slate-700 text-xs font-bold transition-colors">추가</button>
          </form>
          <div className="flex flex-wrap gap-2">
            {paymentMethods.map(pm => (
              <div key={pm.id} className="bg-slate-100 px-3 py-1.5 rounded-full text-xs font-bold text-slate-700 flex items-center gap-2 border border-slate-200">
                {pm.name}
                <button onClick={() => window.confirm('삭제하시겠습니까?') && onDeletePM(pm.id)} className="text-slate-400 hover:text-red-500"><X size={12}/></button>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-yellow-50/50 rounded-xl border border-yellow-200 shadow-sm p-4 sm:p-5">
          <h2 className="text-base font-semibold text-yellow-900 mb-4 flex items-center gap-2"><Target size={18} className="text-yellow-600" /> 저축 목표 관리</h2>
          <form onSubmit={handleAddGoalSubmit} className="space-y-3 mb-4">
            <input type="text" required value={newGoal.name} onChange={e => setNewGoal({...newGoal, name: e.target.value})} placeholder="목표 이름 (예: 유럽 여행)" className="w-full border border-yellow-200 p-2 rounded-lg text-sm focus:ring-2 focus:ring-yellow-400 outline-none bg-white" />
            <div className="relative">
              <input type="text" inputMode="numeric" required value={newGoal.targetAmount} onChange={e => {
                 const val = e.target.value.replace(/[^0-9]/g, '');
                 setNewGoal({...newGoal, targetAmount: val ? new Intl.NumberFormat('ko-KR').format(parseInt(val, 10)) : ''});
              }} placeholder="목표 금액 (숫자만)" className="w-full border border-yellow-200 p-2 pr-8 rounded-lg text-sm focus:ring-2 focus:ring-yellow-400 outline-none text-right bg-white" />
              <span className="absolute right-3 top-2 text-sm font-bold text-slate-400">원</span>
            </div>
            <button type="submit" className="w-full bg-yellow-500 text-yellow-950 py-2.5 rounded-lg hover:bg-yellow-400 text-xs font-bold transition-colors shadow-sm">목표 추가하기</button>
          </form>
          
          <div className="space-y-2">
            {goals.map(goal => (
              <div key={goal.id} className="bg-white p-3 rounded-lg border border-yellow-200 flex justify-between items-center shadow-sm">
                <div>
                  <div className="text-xs font-bold text-slate-800">{goal.name}</div>
                  <div className="text-[11px] text-slate-500 font-medium">목표: {formatCurrency(goal.targetAmount)}</div>
                </div>
                <button onClick={() => window.confirm('이 목표를 삭제하시겠습니까? (기존 내역의 할당 정보는 사라집니다)') && onDeleteGoal(goal.id)} className="p-1.5 text-slate-400 hover:text-red-500 bg-slate-50 rounded transition-colors"><Trash2 size={14}/></button>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden h-fit">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
          <h2 className="text-base font-semibold flex items-center gap-2"><Settings size={18} className="text-slate-500" /> 관리 항목 리스트</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="p-3 font-medium text-center">유형</th>
                <th className="p-3 font-medium">카테고리</th>
                <th className="p-3 font-medium">세부 항목명</th>
                <th className="p-3 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {sortedItems.length === 0 ? (
                <tr><td colSpan="4" className="p-8 text-center text-slate-500">항목이 없습니다.</td></tr>
              ) : (
                sortedItems.map(item => (
                  <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                    {editingItemId === item.id ? (
                      <>
                        <td className="p-2"><select className="border p-1.5 text-xs rounded" value={editFormData.type} onChange={e=>setEditFormData({...editFormData, type: e.target.value})}><option value="INCOME">수입</option><option value="EXPENSE">지출</option><option value="SAVING">저축</option></select></td>
                        <td className="p-2"><input className="border p-1.5 w-full text-xs rounded" value={editFormData.category} onChange={e=>setEditFormData({...editFormData, category: e.target.value})}/></td>
                        <td className="p-2"><input className="border p-1.5 w-full text-xs rounded" value={editFormData.name} onChange={e=>setEditFormData({...editFormData, name: e.target.value})}/></td>
                        <td className="p-2 text-center whitespace-nowrap">
                          <button onClick={() => handleSaveEdit(item.id)} className="p-1.5 text-blue-600 bg-blue-100 rounded mr-1"><Check size={14}/></button>
                          <button onClick={() => setEditingItemId(null)} className="p-1.5 text-slate-600 bg-slate-200 rounded"><X size={14}/></button>
                        </td>
                      </>
                    ) : deletingItemId === item.id ? (
                      <>
                        <td colSpan="3" className="p-3 text-center text-red-600 font-bold text-xs bg-red-50/50">삭제 시 기존 내역은 '삭제된항목'으로 표시됩니다.</td>
                        <td className="p-2 text-center whitespace-nowrap bg-red-50/50">
                          <button onClick={() => { onDeleteItem(item.id); setDeletingItemId(null); }} className="px-2.5 py-1 bg-red-600 text-white text-xs font-bold rounded mr-1 shadow-sm">삭제</button>
                          <button onClick={() => setDeletingItemId(null)} className="px-2.5 py-1 bg-slate-200 text-slate-700 text-xs font-bold rounded shadow-sm">취소</button>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="p-3 text-center">
                          <span className={`px-2 py-1 rounded text-[10px] font-bold ${item.type === 'INCOME' ? 'bg-blue-100 text-blue-700' : item.type === 'EXPENSE' ? 'bg-red-100 text-red-700' : 'bg-emerald-100 text-emerald-700'}`}>
                            {item.type === 'INCOME' ? '수입' : item.type === 'EXPENSE' ? '지출' : '저축'}
                          </span>
                        </td>
                        <td className="p-3 font-bold text-slate-700">{item.category}</td>
                        <td className="p-3 font-medium text-slate-800">{item.name}</td>
                        <td className="p-3 text-center whitespace-nowrap">
                          <button onClick={() => { setEditingItemId(item.id); setEditFormData(item); }} className="p-1.5 text-slate-400 hover:text-blue-600 bg-white border border-slate-200 rounded shadow-sm mr-1"><Pencil size={14}/></button>
                          <button onClick={() => setDeletingItemId(item.id)} className="p-1.5 text-slate-400 hover:text-red-600 bg-white border border-slate-200 rounded shadow-sm"><Trash2 size={14}/></button>
                        </td>
                      </>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [신규 탭] 데이터 백업 / 복원 컴포넌트 ---
const BackupRestoreTab = ({ items, paymentMethods, transactions, quickAdds, goals }) => {
  const fileInputRef = useRef(null);
  const [inlineMessage, setInlineMessage] = useState(null);

  const showMessage = (msg, type = 'success') => {
    setInlineMessage({ msg, type }); setTimeout(() => setInlineMessage(null), 4000);
  };

  const handleExportBackup = () => {
    const backupData = { items, paymentMethods, transactions, quickAdds, goals, exportDate: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url;
    a.download = `가계부_데이터백업_${new Date().toISOString().slice(0,10)}.json`; a.click();
    URL.revokeObjectURL(url); showMessage('데이터 백업 파일이 내 컴퓨터에 다운로드 되었습니다.');
  };

  const handleImportBackup = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const imported = JSON.parse(event.target.result);
        if (imported.items && imported.transactions) {
          saveLocalData('accountbook_local_items', imported.items);
          saveLocalData('accountbook_local_txs', imported.transactions);
          if (imported.paymentMethods) saveLocalData('accountbook_pay_methods', imported.paymentMethods);
          if (imported.quickAdds) saveLocalData('accountbook_quick_adds', imported.quickAdds);
          if (imported.goals) saveLocalData('accountbook_goals', imported.goals);
          showMessage('✅ 데이터 복원 성공! 페이지를 새로고침합니다...', 'success');
          setTimeout(() => window.location.reload(), 1500); 
        } else { showMessage('❌ 올바른 백업 파일이 아닙니다.', 'error'); }
      } catch (err) { showMessage('❌ 파일을 읽는 중 오류가 발생했습니다.', 'error'); }
      e.target.value = '';
    };
    reader.readAsText(file);
  };

  return (
    <div className="max-w-3xl mx-auto mt-4 space-y-6">
      {inlineMessage && (
        <div className={`p-4 text-sm font-bold rounded-xl flex items-center gap-2 shadow-sm ${inlineMessage.type === 'error' ? 'bg-red-50 text-red-600 border border-red-200' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'}`}>
          {inlineMessage.type === 'error' ? <AlertCircle size={18} className="shrink-0"/> : <Check size={18} className="shrink-0"/>}
          {inlineMessage.msg}
        </div>
      )}
      <div className="bg-gradient-to-br from-indigo-50 to-blue-50 rounded-2xl border border-indigo-200 shadow-sm p-6 sm:p-10 relative overflow-hidden">
        <div className="absolute top-1/2 right-4 -translate-y-1/2 p-8 opacity-5 pointer-events-none"><RefreshCw size={200} /></div>
        <h2 className="text-xl font-black text-indigo-900 mb-3 flex items-center gap-2 relative z-10"><Database size={24} className="text-indigo-600" /> 오프라인 데이터 백업 및 복원</h2>
        <p className="text-sm text-slate-600 mb-10 leading-relaxed relative z-10 break-keep">
          이 가계부 앱의 모든 데이터는 안전하게 <strong>현재 사용 중인 기기의 브라우저</strong>에만 저장됩니다.<br/>
          핸드폰을 바꾸거나 다른 컴퓨터에서 내역을 이어가고 싶다면, 데이터를 파일로 내려받아 새 기기에서 불러오세요.
        </p>
        <div className="flex flex-col sm:flex-row gap-4 relative z-10">
          <button onClick={handleExportBackup} className="flex-1 flex items-center justify-center gap-2 bg-white border-2 border-indigo-200 text-indigo-700 py-4 rounded-xl hover:bg-indigo-50 text-sm font-bold transition-all shadow-sm">
            <Download size={20} /> 1. 데이터 백업 파일 저장
          </button>
          <button onClick={() => fileInputRef.current?.click()} className="flex-1 flex items-center justify-center gap-2 bg-indigo-600 text-white py-4 rounded-xl hover:bg-indigo-700 text-sm font-bold transition-all shadow-sm">
            <Upload size={20} /> 2. 백업 파일 불러오기
          </button>
          <input type="file" accept=".json" ref={fileInputRef} style={{ display: 'none' }} onChange={handleImportBackup} />
        </div>
      </div>
    </div>
  );
};

// --- [메인 App 컴포넌트] ---
export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [appName, setAppName] = useState(() => loadLocalData('accountbook_app_name', '나의 가계부'));
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempAppName, setTempAppName] = useState('');

  const [items, setItems] = useState(() => loadLocalData('accountbook_local_items', defaultItems));
  const [transactions, setTransactions] = useState(() => loadLocalData('accountbook_local_txs', []));
  const [budget, setBudget] = useState(() => loadLocalData('accountbook_budget', 1000000));
  const [paymentMethods, setPaymentMethods] = useState(() => loadLocalData('accountbook_pay_methods', defaultPaymentMethods));
  const [quickAdds, setQuickAdds] = useState(() => loadLocalData('accountbook_quick_adds', []));
  const [goals, setGoals] = useState(() => loadLocalData('accountbook_goals', []));

  useEffect(() => saveLocalData('accountbook_local_items', items), [items]);
  useEffect(() => saveLocalData('accountbook_local_txs', transactions), [transactions]);
  useEffect(() => saveLocalData('accountbook_budget', budget), [budget]);
  useEffect(() => saveLocalData('accountbook_pay_methods', paymentMethods), [paymentMethods]);
  useEffect(() => saveLocalData('accountbook_quick_adds', quickAdds), [quickAdds]);
  useEffect(() => saveLocalData('accountbook_goals', goals), [goals]);

  const handleNameSave = () => {
    if (tempAppName.trim() !== '') {
      setAppName(tempAppName.trim());
      saveLocalData('accountbook_app_name', tempAppName.trim());
    }
    setIsEditingName(false);
  };

  const handleAddItem = (newItem) => setItems(prev => [...prev, newItem]);
  const handleUpdateItem = (updatedItem) => setItems(prev => prev.map(i => i.id === updatedItem.id ? updatedItem : i));
  const handleDeleteItem = (id) => setItems(prev => prev.filter(i => i.id !== id));
  
  const handleAddPM = (newPM) => setPaymentMethods(prev => [...prev, newPM]);
  const handleDeletePM = (id) => setPaymentMethods(prev => prev.filter(p => p.id !== id));

  const handleAddTransaction = (newTx) => setTransactions(prev => [...prev, newTx]);
  const handleDeleteTransaction = (id) => setTransactions(prev => prev.filter(tx => tx.id !== id));

  const handleAddGoal = (newGoal) => setGoals(prev => [...prev, newGoal]);
  const handleDeleteGoal = (id) => setGoals(prev => prev.filter(g => g.id !== id));

  const handleImportFixedTransactions = (missingFixedTxs, targetMonthStr) => {
    const newTxs = missingFixedTxs.map(ftx => {
      const originDate = new Date(ftx.date);
      const targetYear = parseInt(targetMonthStr.split('-')[0], 10);
      const targetMonth = parseInt(targetMonthStr.split('-')[1], 10) - 1; 
      let newDateObj = new Date(targetYear, targetMonth, originDate.getDate());

      if (newDateObj.getMonth() !== targetMonth) {
        newDateObj = new Date(targetYear, targetMonth + 1, 0); 
      }
      const dd = String(newDateObj.getDate()).padStart(2, '0');

      return {
        ...ftx,
        id: `t${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
        date: `${targetMonthStr}-${dd}`,
        originalId: ftx.originalId || ftx.id,
        createdAt: new Date().toISOString()
      };
    });
    setTransactions(prev => [...prev, ...newTxs]);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 pb-10">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex flex-col sm:flex-row justify-between w-full md:w-auto gap-4">
            <div className="flex items-center gap-2">
              <div className="bg-slate-800 p-2 rounded-lg shrink-0"><Wallet size={20} className="text-white" /></div>
              {isEditingName ? (
                <input type="text" value={tempAppName} onChange={(e) => setTempAppName(e.target.value)} onBlur={handleNameSave} onKeyDown={(e) => e.key === 'Enter' && handleNameSave()} autoFocus className="text-lg font-bold text-slate-800 border-b-2 border-blue-500 outline-none w-32 sm:w-48 bg-transparent" />
              ) : (
                <h1 className="text-lg font-bold text-slate-800 cursor-pointer hover:text-blue-600 flex items-center gap-2 group transition-colors" onClick={() => { setTempAppName(appName); setIsEditingName(true); }}>
                  {appName} <Pencil size={14} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
                </h1>
              )}
            </div>
          </div>

          <nav className="flex space-x-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0 [&::-webkit-scrollbar]:hidden">
            <button onClick={() => setActiveTab('dashboard')} className={`px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5 shrink-0 transition-all ${activeTab === 'dashboard' ? 'bg-slate-800 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}><LayoutDashboard size={16} />이달의 현황</button>
            <button onClick={() => setActiveTab('transactions')} className={`px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5 shrink-0 transition-all ${activeTab === 'transactions' ? 'bg-slate-800 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}><Receipt size={16} />수입/지출 입력</button>
            <button onClick={() => setActiveTab('calendar')} className={`px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5 shrink-0 transition-all ${activeTab === 'calendar' ? 'bg-slate-800 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}><CalendarDays size={16} />통계/상세조회</button>
            <button onClick={() => setActiveTab('items')} className={`px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5 shrink-0 transition-all ${activeTab === 'items' ? 'bg-slate-800 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}><Settings size={16} />항목 관리</button>
            <button onClick={() => setActiveTab('backup')} className={`px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5 shrink-0 transition-all ${activeTab === 'backup' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-500 hover:bg-indigo-50 hover:text-indigo-600'}`}><Database size={16} />백업/복원</button>
          </nav>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {activeTab === 'dashboard' && <DashboardTab items={items} paymentMethods={paymentMethods} transactions={transactions} onImportFixedTransactions={handleImportFixedTransactions} budget={budget} setBudget={setBudget} onDeleteTransaction={handleDeleteTransaction} goals={goals} />}
        {activeTab === 'transactions' && <TransactionTab items={items} paymentMethods={paymentMethods} transactions={transactions} quickAdds={quickAdds} setQuickAdds={setQuickAdds} goals={goals} onAddTransaction={handleAddTransaction} onDeleteTransaction={handleDeleteTransaction} />}
        {activeTab === 'calendar' && <ReportTab items={items} paymentMethods={paymentMethods} transactions={transactions} onDeleteTransaction={handleDeleteTransaction} goals={goals} />}
        {activeTab === 'items' && <ItemManagementTab items={items} paymentMethods={paymentMethods} goals={goals} onAddItem={handleAddItem} onUpdateItem={handleUpdateItem} onDeleteItem={handleDeleteItem} onAddPM={handleAddPM} onDeletePM={handleDeletePM} onAddGoal={handleAddGoal} onDeleteGoal={handleDeleteGoal} />}
        {activeTab === 'backup' && <BackupRestoreTab items={items} paymentMethods={paymentMethods} transactions={transactions} quickAdds={quickAdds} goals={goals} />}
      </main>
    </div>
  );
}