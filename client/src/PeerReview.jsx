import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Tooltip, Legend
} from 'chart.js';
import { Bar } from 'react-chartjs-2';
import * as XLSX from 'xlsx';

ChartJS.register(CategoryScale, LinearScale, BarElement, Tooltip, Legend);

const PALETTE = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#4a4b4d','#d4c85a','#3a3a3a','#b8cc28','#c2d530','#a8ba20','#6d6e71','#8faa1b'];
const GROUP_NAMES = ['大健康', '半導體', '綠能', '幕僚'];

function detectGroup(filename) {
  for (const g of GROUP_NAMES) {
    if (filename.includes(g)) return g;
  }
  return '未分類';
}

function parsePeerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if (rows.length === 0) { reject('Excel 中沒有資料'); return; }

        const headers = Object.keys(rows[0]);
        const rawName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');
        const group = detectGroup(rawName);
        const batchLabel = rawName;

        const scoreCols = headers.filter(h => h.includes('互評分數'));
        const members = scoreCols.map(col => {
          const match = col.match(/- (.+)$/);
          return match ? match[1].trim() : col;
        });
        const commentCol = headers.find(h => h.includes('想對組員說') || h.includes('說的話')) || '';

        const records = rows.map((row, i) => {
          const scores = {};
          let raterName = '';
          scoreCols.forEach((col, ci) => {
            const val = String(row[col] ?? '').trim();
            if (val === '本人') raterName = members[ci];
            else if (val && !isNaN(Number(val))) scores[members[ci]] = Number(val);
          });
          const comment = commentCol ? String(row[commentCol] ?? '').trim() : '';
          return {
            id: `${Date.now()}-${i}-${Math.random().toString(36).slice(2, 8)}`,
            batch: batchLabel,
            group,
            rater: raterName || `填答者${i + 1}`,
            scores,
            comment: (comment && comment !== '無') ? comment : ''
          };
        });

        const res = await fetch('/api/peers', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ batch: batchLabel, records })
        });
        if (!res.ok) throw (await res.json()).error || '儲存失敗';
        resolve({ count: records.length, batch: batchLabel, group });
      } catch (err) { reject(typeof err === 'string' ? err : err.message); }
    };
    reader.onerror = () => reject('檔案讀取失敗');
    reader.readAsArrayBuffer(file);
  });
}

export default function PeerReview() {
  const [allData, setAllData] = useState([]);
  const [groups, setGroups] = useState([]);
  const [selectedGroup, setSelectedGroup] = useState('');
  const [batches, setBatches] = useState([]);
  const [selectedBatch, setSelectedBatch] = useState('__all__');
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');

  const fetchAll = async () => {
    const res = await fetch('/api/peers');
    const data = res.ok ? await res.json() : [];
    setAllData(data);
    const g = [...new Set(data.map(d => d.group))].sort();
    setGroups(g);
    return { data, groups: g };
  };

  useEffect(() => {
    (async () => {
      const { groups: g } = await fetchAll();
      if (g.length > 0) setSelectedGroup(g[0]);
    })();
  }, []);

  // Update batches when group changes
  useEffect(() => {
    if (!selectedGroup) return;
    const groupData = allData.filter(d => d.group === selectedGroup);
    const b = [...new Set(groupData.map(d => d.batch))].sort();
    setBatches(b);
    setSelectedBatch('__all__');
  }, [selectedGroup, allData]);

  // Filtered data
  const data = allData
    .filter(d => d.group === selectedGroup)
    .filter(d => selectedBatch === '__all__' || d.batch === selectedBatch);

  const handleUpload = async (e) => {
    e.preventDefault();
    const fi = e.target.querySelector('input[type="file"]');
    if (!fi.files[0]) return;
    setUploading(true); setMessage('');
    try {
      const result = await parsePeerExcel(fi.files[0]);
      setMessage(`成功匯入 ${result.count} 筆互評 (${result.group} - ${result.batch})`);
      const { groups: g } = await fetchAll();
      if (result.group && g.includes(result.group)) setSelectedGroup(result.group);
    } catch (err) { setMessage(typeof err === 'string' ? err : '匯入失敗'); }
    setUploading(false); fi.value = '';
  };

  const handleClear = async () => {
    if (!confirm('確定要清除所有互評資料嗎？')) return;
    await fetch('/api/peers', { method: 'DELETE' });
    setAllData([]); setGroups([]); setSelectedGroup(''); setBatches([]);
    setSelectedBatch('__all__'); setMessage('資料已清除');
  };

  // Stats
  const allMembers = [...new Set(data.flatMap(d => Object.keys(d.scores || {})))];
  const memberAvg = {};
  const memberScores = {};
  allMembers.forEach(m => {
    const scores = data.map(d => d.scores?.[m]).filter(s => s !== undefined);
    memberScores[m] = scores;
    memberAvg[m] = scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
  });
  const sortedMembers = [...allMembers].sort((a, b) => memberAvg[b] - memberAvg[a]);
  const comments = data.filter(d => d.comment).map(d => ({ rater: d.rater, batch: d.batch, comment: d.comment }));
  const overallAvg = sortedMembers.length > 0 ? +(Object.values(memberAvg).reduce((a, b) => a + b, 0) / sortedMembers.length).toFixed(2) : 0;

  const avgChart = sortedMembers.length > 0 ? {
    labels: sortedMembers,
    datasets: [{ label: '平均被評分', data: sortedMembers.map(m => memberAvg[m]), backgroundColor: PALETTE.slice(0, sortedMembers.length) }]
  } : null;

  return (
    <div>
      <form className="upload-section" onSubmit={handleUpload}>
        <input type="file" accept=".xlsx,.xls,.csv" aria-label="選擇互評 Excel 檔案" />
        <button className="btn btn-primary" type="submit" disabled={uploading}>{uploading ? '上傳中...' : '匯入互評 Excel'}</button>
        <button className="btn btn-danger" type="button" onClick={handleClear}>清除資料</button>
        {message && <span style={{ color: '#a8ba20', fontWeight: 500 }}>{message}</span>}
      </form>

      {groups.length > 0 && (
        <>
          {/* Group tabs */}
          <div style={{ display: 'flex', gap: 0, marginBottom: 16, borderBottom: '2px solid #e0e0de' }}>
            {groups.map(g => (
              <button key={g} onClick={() => setSelectedGroup(g)} style={{
                padding: '10px 24px', border: 'none', cursor: 'pointer', fontSize: '0.95rem', fontWeight: 500,
                background: selectedGroup === g ? '#fff' : 'transparent',
                color: selectedGroup === g ? '#a8ba20' : '#6d6e71',
                borderBottom: selectedGroup === g ? '3px solid #c2d530' : '3px solid transparent',
                marginBottom: -2, borderRadius: '8px 8px 0 0'
              }}>{g}</button>
            ))}
          </div>

          {/* Batch selector */}
          {batches.length > 0 && (
            <div className="controls">
              <label htmlFor="peer-batch">選擇互評：</label>
              <select id="peer-batch" value={selectedBatch} onChange={e => setSelectedBatch(e.target.value)}>
                <option value="__all__">全部彙總 ({batches.length} 份)</option>
                {batches.map(b => <option key={b} value={b}>{b}</option>)}
              </select>
            </div>
          )}
        </>
      )}

      {data.length > 0 && (
        <>
          <div className="stats-grid">
            <div className="stat-card"><div className="value">{data.length}</div><div className="label">填答人數</div></div>
            <div className="stat-card"><div className="value">{sortedMembers.length}</div><div className="label">被評人數</div></div>
            <div className="stat-card"><div className="value">{overallAvg}</div><div className="label">整體平均分 (滿分5)</div></div>
            <div className="stat-card"><div className="value">{comments.length}</div><div className="label">留言數</div></div>
          </div>

          {avgChart && (
            <div className="chart-card" style={{ marginBottom: 24 }}>
              <h3>📊 各成員平均被評分</h3>
              <Bar data={avgChart} options={{ indexAxis: 'y', scales: { x: { min: 0, max: 5, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } }} />
            </div>
          )}

          <div className="chart-card" style={{ marginBottom: 24, overflowX: 'auto' }}>
            <h3>📋 個人得分明細</h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.9rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #eee', textAlign: 'left' }}>
                  <th style={{ padding: 8 }}>成員</th><th style={{ padding: 8 }}>平均分</th><th style={{ padding: 8 }}>各筆評分</th><th style={{ padding: 8 }}>評價</th>
                </tr>
              </thead>
              <tbody>
                {sortedMembers.map(m => (
                  <tr key={m} style={{ borderBottom: '1px solid #f0f0f0' }}>
                    <td style={{ padding: 8, fontWeight: 500 }}>{m}</td>
                    <td style={{ padding: 8, fontWeight: 600, color: '#a8ba20' }}>{memberAvg[m]}</td>
                    <td style={{ padding: 8, color: '#6d6e71' }}>{memberScores[m].join(', ')}</td>
                    <td style={{ padding: 8, width: '30%' }}>
                      <div style={{ background: '#eef2d0', borderRadius: 4, overflow: 'hidden', height: 8 }}>
                        <div style={{ width: `${(memberAvg[m]/5)*100}%`, height: '100%', background: '#c2d530', borderRadius: 4 }} />
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="chart-card" style={{ marginBottom: 24, overflowX: 'auto' }}>
            <h3>📝 評分矩陣（誰評了誰）</h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #eee' }}>
                  <th style={{ padding: 6, textAlign: 'left' }}>評分者 ↓ / 被評者 →</th>
                  {sortedMembers.map(m => <th key={m} style={{ padding: 6, whiteSpace: 'nowrap', textAlign: 'center' }}>{m}</th>)}
                </tr>
              </thead>
              <tbody>
                {data.map(d => (
                  <tr key={d.id} style={{ borderBottom: '1px solid #f0f0f0' }}>
                    <td style={{ padding: 6, fontWeight: 500, whiteSpace: 'nowrap' }}>
                      {d.rater}
                      {selectedBatch === '__all__' && <span style={{ color: '#999', fontSize: '0.75rem', marginLeft: 4 }}>({d.batch.slice(0, 15)})</span>}
                    </td>
                    {sortedMembers.map(m => {
                      const isSelf = d.rater === m;
                      const score = d.scores?.[m];
                      return (
                        <td key={m} style={{
                          padding: 6, textAlign: 'center',
                          background: isSelf ? '#f5f5f4' : (score ? `rgba(194,213,48,${score/7})` : ''),
                          color: isSelf ? '#999' : '#333', fontWeight: score ? 500 : 400
                        }}>{isSelf ? '本人' : (score ?? '-')}</td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {comments.length > 0 && (
            <div className="chart-card" style={{ marginBottom: 24 }}>
              <h3>💬 給組員的話</h3>
              <div style={{ marginTop: 12 }}>
                {comments.map((c, i) => (
                  <div key={i} style={{ padding: '10px 0', borderBottom: '1px solid #f0f0f0', display: 'flex', gap: 8, alignItems: 'baseline' }}>
                    <span style={{ color: '#6d6e71', fontSize: '0.85rem', minWidth: 60, flexShrink: 0 }}>{c.rater}</span>
                    {selectedBatch === '__all__' && <span style={{ background: '#eef2d0', padding: '1px 6px', borderRadius: 4, fontSize: '0.75rem', color: '#a8ba20', flexShrink: 0 }}>{c.batch.slice(0, 15)}</span>}
                    <span style={{ fontWeight: 500 }}>{c.comment}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </>
      )}

      {groups.length === 0 && (
        <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
          <p style={{ fontSize: '1.2rem' }}>尚無互評資料，請先匯入 Excel 檔案</p>
          <p style={{ marginTop: 8, fontSize: '0.9rem' }}>檔名包含組別名稱（大健康、半導體、綠能、幕僚）可自動分類</p>
        </div>
      )}

      {groups.length > 0 && data.length === 0 && (
        <div style={{ textAlign: 'center', padding: 40, color: '#999' }}>
          <p>此組別尚無互評資料</p>
        </div>
      )}
    </div>
  );
}
