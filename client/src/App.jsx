import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS, ArcElement, Tooltip, Legend,
  CategoryScale, LinearScale, BarElement, RadialLinearScale, PointElement, LineElement, Filler
} from 'chart.js';
import { Pie, Bar, Radar } from 'react-chartjs-2';
import { parseExcelFile, getBatches, getFeedbacks, getStats, clearData } from './storage.js';
import PeerReview from './PeerReview.jsx';
import DeleteModal from './DeleteModal.jsx';

ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement, RadialLinearScale, PointElement, LineElement, Filler);

const COLORS_SAT = {
  '非常滿意': '#a8ba20', '滿意': '#c2d530', '普通': '#d4c85a', '不滿意': '#e0a050', '非常不滿意': '#c45040'
};
const PALETTE = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#4a4b4d','#d4c85a','#3a3a3a','#b8cc28'];

function shortLabel(col) {
  return col.replace('課程安排 - ', '').replace('講師授課情形 - ', '').replace('課程難易度 - ', '難易度: ');
}

export default function App() {
  const [page, setPage] = useState('feedback');
  const [batches, setBatches] = useState([]);
  const [selectedBatch, setSelectedBatch] = useState('');
  const [selectedGroup, setSelectedGroup] = useState('');
  const [groups, setGroups] = useState([]);
  const [stats, setStats] = useState(null);
  const [feedbacks, setFeedbacks] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');
  const [selectedExtraField, setSelectedExtraField] = useState('');
  const [showDeleteModal, setShowDeleteModal] = useState(false);

  const refresh = async (batch, group) => {
    const b = batch || selectedBatch;
    const g = group !== undefined ? group : selectedGroup;
    const s = await getStats(b, g);
    setStats(s);
    const f = await getFeedbacks(b, g);
    setFeedbacks(f);
    if (!g && s.groups) setGroups(Object.keys(s.groups));
  };

  const refreshBatches = async () => {
    const b = await getBatches();
    setBatches(b);
    return b;
  };

  useEffect(() => {
    (async () => {
      const b = await refreshBatches();
      if (b.length > 0) { setSelectedBatch(b[0]); await refresh(b[0], ''); }
    })();
  }, []);

  useEffect(() => { if (selectedBatch) refresh(selectedBatch, selectedGroup); }, [selectedBatch, selectedGroup]);

  const handleUpload = async (e) => {
    e.preventDefault();
    const fi = e.target.querySelector('input[type="file"]');
    if (!fi.files[0]) return;
    setUploading(true); setMessage('');
    try {
      const result = await parseExcelFile(fi.files[0]);
      setMessage(`成功匯入 ${result.count} 筆回饋 (${result.batch})`);
      const nb = await refreshBatches();
      if (nb.length > 0) { setSelectedGroup(''); setSelectedBatch(nb[nb.length - 1]); await refresh(nb[nb.length - 1], ''); }
    } catch (err) { setMessage(typeof err === 'string' ? err : '匯入失敗'); }
    setUploading(false); fi.value = '';
  };

  const handleClear = async (batchesToDelete) => {
    for (const b of batchesToDelete) {
      await fetch(`/api/feedbacks?batch=${encodeURIComponent(b)}`, { method: 'DELETE' });
    }
    const nb = await refreshBatches();
    if (nb.length > 0) {
      setSelectedGroup(''); setSelectedBatch(nb[0]); await refresh(nb[0], '');
    } else {
      setBatches([]); setSelectedBatch(''); setSelectedGroup(''); setGroups([]);
      setStats(null); setFeedbacks([]);
    }
    setMessage(`已刪除 ${batchesToDelete.length} 筆資料`);
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
      backgroundColor: 'rgba(194,213,48,0.2)', borderColor: '#a8ba20', pointBackgroundColor: '#a8ba20'
    }]
  } : null;

  const allRatingKeys = feedbacks.length > 0
    ? [...new Set(feedbacks.flatMap(f => Object.keys(f.ratings || {})))]
    : [];

  return (
    <div className="app">
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h1>📊 崇越雲拓計畫課程回饋分析</h1>
        <div style={{ display: 'flex', gap: 4, background: '#e0e0de', borderRadius: 8, padding: 3 }}>
          <button
            onClick={() => setPage('feedback')}
            style={{
              padding: '8px 16px', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: '0.9rem', fontWeight: 500,
              background: page === 'feedback' ? '#c2d530' : 'transparent',
              color: page === 'feedback' ? '#fff' : '#6d6e71'
            }}
          >課程回饋</button>
          <button
            onClick={() => setPage('peer')}
            style={{
              padding: '8px 16px', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: '0.9rem', fontWeight: 500,
              background: page === 'peer' ? '#c2d530' : 'transparent',
              color: page === 'peer' ? '#fff' : '#6d6e71'
            }}
          >各組互評</button>
        </div>
      </header>

      {page === 'peer' ? <PeerReview /> : (
      <>

      <form className="upload-section" onSubmit={handleUpload}>
        <input type="file" accept=".xlsx,.xls,.csv" aria-label="選擇 Excel 檔案" />
        <button className="btn btn-primary" type="submit" disabled={uploading}>{uploading ? '上傳中...' : '匯入 Excel'}</button>
        <button className="btn btn-danger" type="button" onClick={() => setShowDeleteModal(true)}>清除資料</button>
        {message && <span style={{ color: '#a8ba20', fontWeight: 500 }}>{message}</span>}
      </form>

      {batches.length > 0 && (
        <div className="controls">
          <label htmlFor="batch-select">選擇課程：</label>
          <select id="batch-select" value={selectedBatch} onChange={e => { setSelectedGroup(''); setSelectedExtraField(''); setSelectedBatch(e.target.value); }}>
            {batches.map(b => <option key={b} value={b}>{b}</option>)}
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
                        <div style={{ background: '#eef2d0', borderRadius: 4, overflow: 'hidden', height: 8 }}>
                          <div style={{ width: `${(stats.ratingAvg[k]/5)*100}%`, height: '100%', background: i < 4 ? '#c2d530' : '#a8ba20', borderRadius: 4 }} />
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
        <div className="chart-card" style={{ marginTop: 24, maxHeight: 300, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
          <h3 style={{ flexShrink: 0 }}>💬 建議回饋</h3>
          <div style={{ marginTop: 8, overflowY: 'auto', fontSize: '0.85rem' }}>
            {stats.suggestions.map((s, i) => (
              <div key={i} style={{ padding: '12px 0', borderBottom: '1px solid #f0f0f0' }}>
                <div style={{ display: 'flex', gap: 8, marginBottom: 4 }}>
                  <span style={{ background: '#eef2d0', padding: '2px 8px', borderRadius: 4, fontSize: '0.85rem', color: '#a8ba20' }}>{s.type}</span>
                  <span style={{ color: '#888', fontSize: '0.85rem' }}>{s.name}</span>
                </div>
                <div style={{ color: '#333' }}>{s.content}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {stats?.extras?.length > 0 && (() => {
        const extraFields = [...new Set(stats.extras.map(e => e.field))];
        const filtered = selectedExtraField ? stats.extras.filter(e => e.field === selectedExtraField) : stats.extras;
        return (
          <div className="chart-card" style={{ marginTop: 24 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 12 }}>
              <h3 style={{ margin: 0 }}>📝 其他回饋</h3>
              <select
                value={selectedExtraField}
                onChange={e => setSelectedExtraField(e.target.value)}
                style={{ padding: '6px 10px', border: '1px solid #e0e0de', borderRadius: 8, fontSize: '0.9rem' }}
              >
                <option value="">全部欄位</option>
                {extraFields.map(f => <option key={f} value={f}>{shortLabel(f)}</option>)}
              </select>
              <span style={{ color: '#6d6e71', fontSize: '0.85rem' }}>{filtered.length} 筆</span>
            </div>
            <div>
              {filtered.map((e, i) => (
                <div key={i} style={{ padding: '10px 0', borderBottom: '1px solid #f0f0f0', display: 'flex', gap: 10, alignItems: 'baseline' }}>
                  <span style={{ color: '#6d6e71', fontSize: '0.85rem', minWidth: 70, flexShrink: 0 }}>{e.name}</span>
                  {!selectedExtraField && <span style={{ background: '#eef2d0', padding: '2px 8px', borderRadius: 4, fontSize: '0.8rem', color: '#a8ba20', flexShrink: 0 }}>{shortLabel(e.field)}</span>}
                  <span style={{ fontWeight: 500 }}>{e.value}</span>
                </div>
              ))}
            </div>
          </div>
        );
      })()}

      {feedbacks.length > 0 && (
        <div style={{ marginTop: 24 }}>
          <div className="chart-card" style={{ borderRadius: '12px 12px 0 0', paddingBottom: 0 }}>
            <h3>📋 個人填答明細</h3>
          </div>
          <div style={{ background: '#fff', borderRadius: '0 0 12px 12px', boxShadow: '0 2px 8px rgba(0,0,0,0.06)', maxHeight: 480, overflowY: 'auto', overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '0.9rem' }}>
            <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
              <tr>
                <th style={{ padding: 6, whiteSpace: 'nowrap', background: '#fff', borderBottom: '2px solid #e0e0de', textAlign: 'left' }}>姓名</th>
                <th style={{ padding: 6, whiteSpace: 'nowrap', background: '#fff', borderBottom: '2px solid #e0e0de', textAlign: 'left' }}>組別</th>
                {allRatingKeys.map(k => <th key={k} style={{ padding: 6, whiteSpace: 'nowrap', background: '#fff', borderBottom: '2px solid #e0e0de', textAlign: 'left' }}>{shortLabel(k)}</th>)}
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
        </div>
      )}

      {batches.length === 0 && (
        <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
          <p style={{ fontSize: '1.2rem' }}>尚無資料，請先匯入 Excel 檔案</p>
          <p style={{ marginTop: 8, fontSize: '0.9rem' }}>支援自動偵測欄位格式，資料儲存在瀏覽器中</p>
        </div>
      )}
      </>
      )}

      {showDeleteModal && (
        <DeleteModal
          title="選擇要刪除的課程回饋"
          items={batches}
          onDelete={handleClear}
          onClose={() => setShowDeleteModal(false)}
        />
      )}
    </div>
  );
}
