import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS, ArcElement, Tooltip, Legend,
  CategoryScale, LinearScale, BarElement, RadialLinearScale, PointElement, LineElement, Filler
} from 'chart.js';
import { Pie, Bar, Radar } from 'react-chartjs-2';
import { parseExcelFile, getDates, getFeedbacks, getStats, clearData } from './storage.js';

ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement, RadialLinearScale, PointElement, LineElement, Filler);

const COLORS_SAT = {
  '非常滿意': '#2dc653', '滿意': '#4361ee', '普通': '#f4a261', '不滿意': '#e76f51', '非常不滿意': '#e63946'
};
const PALETTE = ['#4361ee','#7209b7','#f72585','#4cc9f0','#3a0ca3','#560bad','#e63946','#2dc653'];

function shortLabel(col) {
  return col.replace('課程安排 - ', '').replace('講師授課情形 - ', '').replace('課程難易度 - ', '難易度: ');
}

export default function App() {
  const [dates, setDates] = useState([]);
  const [selectedDate, setSelectedDate] = useState('');
  const [selectedGroup, setSelectedGroup] = useState('');
  const [groups, setGroups] = useState([]);
  const [stats, setStats] = useState(null);
  const [feedbacks, setFeedbacks] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');

  const refresh = (date, group) => {
    const d = date || selectedDate;
    const g = group !== undefined ? group : selectedGroup;
    const s = getStats(d, g);
    setStats(s);
    setFeedbacks(getFeedbacks(d, g));
    if (!g && s.groups) setGroups(Object.keys(s.groups));
  };

  const refreshDates = () => {
    const d = getDates();
    setDates(d);
    return d;
  };

  useEffect(() => {
    const d = refreshDates();
    if (d.length > 0) { setSelectedDate(d[0]); refresh(d[0], ''); }
  }, []);

  useEffect(() => { if (selectedDate) refresh(selectedDate, selectedGroup); }, [selectedDate, selectedGroup]);

  const handleUpload = async (e) => {
    e.preventDefault();
    const fi = e.target.querySelector('input[type="file"]');
    if (!fi.files[0]) return;
    setUploading(true); setMessage('');
    try {
      const result = await parseExcelFile(fi.files[0]);
      setMessage(`成功匯入 ${result.count} 筆回饋 (${result.date})`);
      const nd = refreshDates();
      if (nd.length > 0) { setSelectedGroup(''); setSelectedDate(nd[0]); refresh(nd[0], ''); }
    } catch (err) { setMessage(typeof err === 'string' ? err : '匯入失敗'); }
    setUploading(false); fi.value = '';
  };

  const handleClear = () => {
    if (!confirm('確定要清除所有資料嗎？')) return;
    clearData();
    setDates([]); setSelectedDate(''); setSelectedGroup(''); setGroups([]);
    setStats(null); setFeedbacks([]); setMessage('資料已清除');
  };

  const ratingKeys = stats?.ratingAvg ? Object.keys(stats.ratingAvg) : [];

  const distChart = stats?.distribution ? {
    labels: Object.keys(stats.distribution),
    datasets: [{ data: Object.values(stats.distribution), backgroundColor: Object.keys(stats.distribution).map(k => COLORS_SAT[k] || '#888') }]
  } : null;

  const groupChart = stats?.groups && Object.keys(stats.groups).length > 0 ? {
    labels: Object.keys(stats.groups),
    datasets: [{ label: '人數', data: Object.values(stats.groups), backgroundColor: PALETTE }]
  } : null;

  const radarData = ratingKeys.length > 0 ? {
    labels: ratingKeys.map(shortLabel),
    datasets: [{
      label: '平均分數', data: ratingKeys.map(k => stats.ratingAvg[k]),
      backgroundColor: 'rgba(67,97,238,0.2)', borderColor: '#4361ee', pointBackgroundColor: '#4361ee'
    }]
  } : null;

  const allRatingKeys = feedbacks.length > 0
    ? [...new Set(feedbacks.flatMap(f => Object.keys(f.ratings || {})))]
    : [];

  return (
    <div className="app">
      <header>
        <h1>📊 課程回饋分析系統</h1>
        <p>匯入課程回饋 Excel，依日期查看分析結果</p>
      </header>

      <form className="upload-section" onSubmit={handleUpload}>
        <input type="file" accept=".xlsx,.xls,.csv" aria-label="選擇 Excel 檔案" />
        <button className="btn btn-primary" type="submit" disabled={uploading}>{uploading ? '上傳中...' : '匯入 Excel'}</button>
        <button className="btn btn-danger" type="button" onClick={handleClear}>清除資料</button>
        {message && <span style={{ color: '#4361ee', fontWeight: 500 }}>{message}</span>}
      </form>

      {dates.length > 0 && (
        <div className="controls">
          <label htmlFor="date-select">選擇日期：</label>
          <select id="date-select" value={selectedDate} onChange={e => { setSelectedGroup(''); setSelectedDate(e.target.value); }}>
            {dates.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <label htmlFor="group-select">篩選組別：</label>
          <select id="group-select" value={selectedGroup} onChange={e => setSelectedGroup(e.target.value)}>
            <option value="">全部組別</option>
            {groups.map(g => <option key={g} value={g}>{g}</option>)}
          </select>
        </div>
      )}

      {stats && stats.total > 0 && (
        <>
          <div className="stats-grid">
            <div className="stat-card"><div className="value">{stats.total}</div><div className="label">填答人數</div></div>
            <div className="stat-card"><div className="value">{stats.overallAvg}</div><div className="label">整體平均分 (滿分5)</div></div>
            <div className="stat-card"><div className="value">{Object.keys(stats.groups).length}</div><div className="label">參加組別數</div></div>
            <div className="stat-card"><div className="value">{stats.suggestions?.length || 0}</div><div className="label">建議回饋數</div></div>
          </div>

          <div className="charts-row">
            {radarData && (
              <div className="chart-card">
                <h3>各項目平均分數</h3>
                <Radar data={radarData} options={{ scales: { r: { min: 0, max: 5, ticks: { stepSize: 1 } } } }} />
              </div>
            )}
            {distChart && (
              <div className="chart-card">
                <h3>整體滿意度分佈</h3>
                <Pie data={distChart} />
              </div>
            )}
          </div>

          <div className="charts-row">
            {groupChart && (
              <div className="chart-card">
                <h3>各組別填答人數</h3>
                <Bar data={groupChart} options={{ plugins: { legend: { display: false } } }} />
              </div>
            )}
            <div className="chart-card">
              <h3>各項目平均分數明細</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr style={{ borderBottom: '2px solid #eee', textAlign: 'left' }}>
                    <th style={{ padding: 8 }}>項目</th><th style={{ padding: 8 }}>平均分</th><th style={{ padding: 8 }}>評價</th>
                  </tr>
                </thead>
                <tbody>
                  {ratingKeys.map((k, i) => (
                    <tr key={k} style={{ borderBottom: '1px solid #f0f0f0' }}>
                      <td style={{ padding: 8 }}>{shortLabel(k)}</td>
                      <td style={{ padding: 8, fontWeight: 600 }}>{stats.ratingAvg[k]}</td>
                      <td style={{ padding: 8 }}>
                        <div style={{ background: '#e8edff', borderRadius: 4, overflow: 'hidden', height: 8 }}>
                          <div style={{ width: `${(stats.ratingAvg[k]/5)*100}%`, height: '100%', background: i < 4 ? '#4361ee' : '#7209b7', borderRadius: 4 }} />
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {stats?.suggestions?.length > 0 && (
        <div className="chart-card" style={{ marginTop: 24 }}>
          <h3>💬 建議回饋</h3>
          <div style={{ marginTop: 12 }}>
            {stats.suggestions.map((s, i) => (
              <div key={i} style={{ padding: '12px 0', borderBottom: '1px solid #f0f0f0' }}>
                <div style={{ display: 'flex', gap: 8, marginBottom: 4 }}>
                  <span style={{ background: '#e8edff', padding: '2px 8px', borderRadius: 4, fontSize: '0.85rem', color: '#4361ee' }}>{s.type}</span>
                  <span style={{ color: '#888', fontSize: '0.85rem' }}>{s.name}</span>
                </div>
                <div style={{ color: '#333' }}>{s.content}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {stats?.extras?.length > 0 && (
        <div className="chart-card" style={{ marginTop: 24 }}>
          <h3>📝 其他回饋（應用場景、學習方向等）</h3>
          <div style={{ marginTop: 12 }}>
            {stats.extras.map((e, i) => (
              <div key={i} style={{ padding: '8px 0', borderBottom: '1px solid #f0f0f0', display: 'flex', gap: 8, alignItems: 'baseline' }}>
                <span style={{ color: '#888', fontSize: '0.85rem', minWidth: 60 }}>{e.name}</span>
                <span style={{ background: '#f0f2f5', padding: '2px 8px', borderRadius: 4, fontSize: '0.8rem', color: '#666' }}>{shortLabel(e.field)}</span>
                <span>{e.value}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {feedbacks.length > 0 && (
        <div className="chart-card" style={{ marginTop: 24, overflowX: 'auto' }}>
          <h3>📋 個人填答明細</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.9rem' }}>
            <thead>
              <tr style={{ borderBottom: '2px solid #eee', textAlign: 'left' }}>
                <th style={{ padding: 6, whiteSpace: 'nowrap' }}>姓名</th>
                <th style={{ padding: 6, whiteSpace: 'nowrap' }}>組別</th>
                {allRatingKeys.map(k => <th key={k} style={{ padding: 6, whiteSpace: 'nowrap' }}>{shortLabel(k)}</th>)}
              </tr>
            </thead>
            <tbody>
              {feedbacks.map(f => (
                <tr key={f.id} style={{ borderBottom: '1px solid #f0f0f0' }}>
                  <td style={{ padding: 6 }}>{f.name}</td>
                  <td style={{ padding: 6 }}>{f.group}</td>
                  {allRatingKeys.map(k => {
                    const r = f.ratings?.[k];
                    return <td key={k} style={{ padding: 6, color: COLORS_SAT[r?.label] || '#333', fontWeight: 500 }}>{r?.label || '-'}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {dates.length === 0 && (
        <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
          <p style={{ fontSize: '1.2rem' }}>尚無資料，請先匯入 Excel 檔案</p>
          <p style={{ marginTop: 8, fontSize: '0.9rem' }}>支援自動偵測欄位格式，資料儲存在瀏覽器中</p>
        </div>
      )}
    </div>
  );
}
