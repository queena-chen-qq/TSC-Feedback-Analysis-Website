import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, PointElement, LineElement, Tooltip, Legend
} from 'chart.js';
import { Bar, Line } from 'react-chartjs-2';
import * as XLSX from 'xlsx';

ChartJS.register(CategoryScale, LinearScale, BarElement, PointElement, LineElement, Tooltip, Legend);

const PALETTE = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#4a4b4d','#d4c85a','#3a3a3a','#b8cc28','#e0a050','#c45040','#7cb342','#558b2f'];
const LINE_COLORS = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#e0a050','#c45040','#4a4b4d','#7cb342','#558b2f','#d4c85a','#b8cc28','#3a3a3a'];
const GROUP_KEYWORDS = ['大健康', '半導體', '綠能', '幕僚'];

function detectGroup(filename) {
  for (const g of GROUP_KEYWORDS) {
    if (filename.includes(g)) return g;
  }
  return filename;
}

// Read raw rows from a sheet, find the real header row by keyword, return { headers, dataRows }
function extractTable(sheet, keyword) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  let headerIdx = -1;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    if (raw[i] && raw[i].some(c => String(c).includes(keyword))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) return { headers: [], dataRows: [] };

  const headers = raw[headerIdx].map(h => String(h ?? '').trim());
  const dataRows = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row || row.every(c => c === '' || c === undefined || c === null)) continue;
    const obj = {};
    headers.forEach((h, ci) => { if (h) obj[h] = row[ci]; });
    dataRows.push(obj);
  }
  return { headers, dataRows };
}

function parseGroupExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const group = detectGroup(file.name.replace(/\.(xlsx|xls|csv)$/i, ''));

        // --- Parse overview (first sheet) ---
        const { headers: ovH, dataRows: ovRows } = extractTable(wb.Sheets[wb.SheetNames[0]], '組員姓名');
        const nameCol = ovH.find(h => h.includes('姓名')) || ovH[0];
        const avgCol = ovH.find(h => h.includes('總平均')) || '';
        const sessionCols = ovH.filter(h => h && h !== nameCol && h !== avgCol);

        const overview = ovRows
          .filter(r => String(r[nameCol] ?? '').trim())
          .map(r => {
            const name = String(r[nameCol]).trim();
            const sessions = {};
            sessionCols.forEach(col => {
              const v = parseFloat(r[col]);
              if (!isNaN(v)) sessions[col] = v;
            });
            const avg = avgCol ? parseFloat(r[avgCol]) : 0;
            return { name, sessions, totalAvg: isNaN(avg) ? 0 : avg };
          });

        // --- Parse each session sheet (skip first) ---
        const sessionSheets = [];
        for (let i = 1; i < wb.SheetNames.length; i++) {
          const sheetName = wb.SheetNames[i];
          const { headers, dataRows } = extractTable(wb.Sheets[sheetName], '評分者');
          if (dataRows.length === 0) continue;

          const raterCol = headers.find(h => h.includes('評分者')) || headers[0];
          const commentCol = headers.find(h => h.includes('留言') || h.includes('想對') || h.includes('說的話')) || '';
          const avgRow = headers.find(h => h.includes('平均'));
          const members = headers.filter(h => h && h !== raterCol && h !== commentCol && !h.includes('平均'));

          const records = dataRows
            .filter(r => {
              const rater = String(r[raterCol] ?? '').trim();
              return rater && !rater.includes('平均');
            })
            .map(r => {
              const rater = String(r[raterCol]).trim();
              const scores = {};
              members.forEach(m => {
                const val = String(r[m] ?? '').trim();
                if (val === '本人') { /* self */ }
                else if (val && val !== '-' && val !== '' && !isNaN(Number(val))) {
                  scores[m] = Number(val);
                }
              });
              const comment = commentCol ? String(r[commentCol] ?? '').trim() : '';
              return { rater, scores, comment: (comment && comment !== '無') ? comment : '' };
            });

          sessionSheets.push({ name: sheetName, members, records });
        }

        resolve({ group, overview, sessions: sessionSheets });
      } catch (err) { reject(err.message); }
    };
    reader.onerror = () => reject('檔案讀取失敗');
    reader.readAsArrayBuffer(file);
  });
}

// ─── Component ───
export default function PeerReview() {
  const [groups, setGroups] = useState({});       // { groupName: { overview, sessions } }
  const [selectedGroup, setSelectedGroup] = useState('');
  const [selectedTab, setSelectedTab] = useState('overview');
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');

  // Load from DynamoDB
  const fetchAll = async () => {
    try {
      const res = await fetch('/api/peers');
      if (!res.ok) return;
      const items = await res.json();
      const g = {};
      items.forEach(item => {
        if (item.group && item.groupData) g[item.group] = item.groupData;
      });
      setGroups(g);
      const names = Object.keys(g).sort();
      if (names.length > 0 && !selectedGroup) setSelectedGroup(names[0]);
    } catch {}
  };

  useEffect(() => { fetchAll(); }, []);

  const handleUpload = async (e) => {
    e.preventDefault();
    const fi = e.target.querySelector('input[type="file"]');
    if (!fi.files[0]) return;
    setUploading(true); setMessage('');
    try {
      const result = await parseGroupExcel(fi.files[0]);
      const record = {
        id: `peer-${Date.now()}`,
        batch: result.group,
        group: result.group,
        groupData: { overview: result.overview, sessions: result.sessions }
      };
      const res = await fetch('/api/peers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ batch: result.group, records: [record] })
      });
      if (!res.ok) throw '儲存失敗';
      setGroups(prev => ({ ...prev, [result.group]: record.groupData }));
      setSelectedGroup(result.group);
      setSelectedTab('overview');
      setMessage(`成功匯入 ${result.group} 互評資料（${result.sessions.length} 堂課）`);
    } catch (err) { setMessage(typeof err === 'string' ? err : '匯入失敗'); }
    setUploading(false); fi.value = '';
  };

  const handleClear = async () => {
    if (!confirm('確定要清除所有互評資料嗎？')) return;
    await fetch('/api/peers', { method: 'DELETE' });
    setGroups({}); setSelectedGroup(''); setSelectedTab('overview'); setMessage('資料已清除');
  };

  const groupNames = Object.keys(groups).sort();
  const cur = groups[selectedGroup];
  const overview = cur?.overview || [];
  const sessions = cur?.sessions || [];
  const curSession = sessions.find(s => s.name === selectedTab);

  // ─── Overview charts ───
  const sessionLabels = overview.length > 0 ? Object.keys(overview[0].sessions) : [];
  const sorted = [...overview].sort((a, b) => b.totalAvg - a.totalAvg);

  const barData = sorted.length > 0 ? {
    labels: sorted.map(m => m.name),
    datasets: [{ label: '總平均', data: sorted.map(m => m.totalAvg), backgroundColor: PALETTE }]
  } : null;

  const lineData = overview.length > 0 && sessionLabels.length > 1 ? {
    labels: sessionLabels.map(l => l.replace(/第\d+堂 \(/, '').replace(')', '').replace(/第\d+堂/, '')),
    datasets: overview.map((m, i) => ({
      label: m.name,
      data: sessionLabels.map(l => m.sessions[l] ?? null),
      borderColor: LINE_COLORS[i % LINE_COLORS.length],
      backgroundColor: 'transparent',
      tension: 0.3, pointRadius: 4, spanGaps: true
    }))
  } : null;

  // ─── Render ───
  return (
    <div>
      {/* Upload */}
      <form className="upload-section" onSubmit={handleUpload}>
        <input type="file" accept=".xlsx,.xls" aria-label="選擇互評 Excel" />
        <button className="btn btn-primary" type="submit" disabled={uploading}>
          {uploading ? '上傳中...' : '匯入互評 Excel'}
        </button>
        <button className="btn btn-danger" type="button" onClick={handleClear}>清除資料</button>
        {message && <span style={{ color: '#a8ba20', fontWeight: 500 }}>{message}</span>}
      </form>

      {/* Empty state */}
      {groupNames.length === 0 && (
        <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
          <p style={{ fontSize: '1.2rem' }}>尚無互評資料</p>
          <p style={{ marginTop: 8, fontSize: '0.9rem' }}>每組上傳一份 Excel（檔名含組別：大健康、半導體、綠能、幕僚）</p>
          <p style={{ fontSize: '0.9rem' }}>Excel 內含總覽 sheet + 各堂課互評 sheet</p>
        </div>
      )}

      {groupNames.length > 0 && (
        <>
          {/* ── Group tabs ── */}
          <div style={{ display: 'flex', gap: 0, borderBottom: '2px solid #e0e0de', marginBottom: 16 }}>
            {groupNames.map(g => (
              <button key={g} onClick={() => { setSelectedGroup(g); setSelectedTab('overview'); }} style={{
                padding: '10px 24px', border: 'none', cursor: 'pointer', fontSize: '0.95rem', fontWeight: 600,
                background: selectedGroup === g ? '#fff' : 'transparent',
                color: selectedGroup === g ? '#a8ba20' : '#6d6e71',
                borderBottom: selectedGroup === g ? '3px solid #c2d530' : '3px solid transparent',
                marginBottom: -2, borderRadius: '8px 8px 0 0',
                transition: 'all 0.15s'
              }}>{g}</button>
            ))}
          </div>

          {/* ── Session tabs ── */}
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 20 }}>
            <button onClick={() => setSelectedTab('overview')} className="btn" style={{
              padding: '7px 16px', fontSize: '0.85rem',
              background: selectedTab === 'overview' ? '#c2d530' : '#e8e8e6',
              color: selectedTab === 'overview' ? '#fff' : '#6d6e71'
            }}>📊 總覽</button>
            {sessions.map(s => (
              <button key={s.name} onClick={() => setSelectedTab(s.name)} className="btn" style={{
                padding: '7px 16px', fontSize: '0.85rem',
                background: selectedTab === s.name ? '#c2d530' : '#e8e8e6',
                color: selectedTab === s.name ? '#fff' : '#6d6e71'
              }}>{s.name}</button>
            ))}
          </div>

          {/* ══════ OVERVIEW ══════ */}
          {selectedTab === 'overview' && overview.length > 0 && (
            <>
              {/* Stats */}
              <div className="stats-grid">
                <div className="stat-card"><div className="value">{overview.length}</div><div className="label">組員人數</div></div>
                <div className="stat-card"><div className="value">{sessionLabels.length}</div><div className="label">互評次數</div></div>
                <div className="stat-card">
                  <div className="value">{sorted.length > 0 ? sorted[0].totalAvg.toFixed(2) : '-'}</div>
                  <div className="label">最高總平均</div>
                </div>
                <div className="stat-card">
                  <div className="value">{(overview.reduce((s, m) => s + m.totalAvg, 0) / overview.length).toFixed(2)}</div>
                  <div className="label">全組平均</div>
                </div>
              </div>

              {/* Bar: total avg ranking */}
              {barData && (
                <div className="chart-card" style={{ marginBottom: 24 }}>
                  <h3>📊 各成員總平均排名</h3>
                  <Bar data={barData} options={{
                    indexAxis: 'y',
                    scales: { x: { min: 0, max: 5, ticks: { stepSize: 1 } } },
                    plugins: { legend: { display: false } }
                  }} />
                </div>
              )}

              {/* Line: trend */}
              {lineData && (
                <div className="chart-card" style={{ marginBottom: 24 }}>
                  <h3>📈 歷次分數趨勢</h3>
                  <Line data={lineData} options={{
                    scales: { y: { min: 0, max: 5, ticks: { stepSize: 1 } } },
                    plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 8, font: { size: 11 } } } }
                  }} />
                </div>
              )}

              {/* Table: overview */}
              <div className="chart-card" style={{ overflowX: 'auto' }}>
                <h3>📋 各堂課平均分數</h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.9rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid #e0e0de', textAlign: 'center' }}>
                      <th style={{ padding: 8, textAlign: 'left' }}>組員</th>
                      {sessionLabels.map(l => <th key={l} style={{ padding: 8, whiteSpace: 'nowrap' }}>{l}</th>)}
                      <th style={{ padding: 8, color: '#a8ba20', fontWeight: 700 }}>總平均</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sorted.map(m => (
                      <tr key={m.name} style={{ borderBottom: '1px solid #f0f0f0' }}>
                        <td style={{ padding: 8, fontWeight: 500 }}>{m.name}</td>
                        {sessionLabels.map(l => {
                          const v = m.sessions[l];
                          const color = v >= 4.5 ? '#a8ba20' : v < 3 ? '#c45040' : v < 3.5 ? '#e0a050' : '#333';
                          return <td key={l} style={{ padding: 8, textAlign: 'center', color, fontWeight: v >= 4.5 || v < 3 ? 600 : 400 }}>{v != null ? v.toFixed(2) : '-'}</td>;
                        })}
                        <td style={{ padding: 8, textAlign: 'center', fontWeight: 700, color: '#a8ba20', fontSize: '1rem' }}>{m.totalAvg.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {/* ══════ SESSION DETAIL ══════ */}
          {curSession && selectedTab !== 'overview' && (() => {
            const { members, records } = curSession;
            const comments = records.filter(r => r.comment);

            const memberAvg = {};
            members.forEach(m => {
              const scores = records.map(r => r.scores[m]).filter(s => s !== undefined);
              memberAvg[m] = scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
            });
            const sortedM = [...members].sort((a, b) => memberAvg[b] - memberAvg[a]);

            const sessionBar = {
              labels: sortedM,
              datasets: [{ label: '平均被評分', data: sortedM.map(m => memberAvg[m]), backgroundColor: PALETTE }]
            };

            return (
              <>
                {/* Stats */}
                <div className="stats-grid">
                  <div className="stat-card"><div className="value">{records.length}</div><div className="label">填答人數</div></div>
                  <div className="stat-card"><div className="value">{members.length}</div><div className="label">被評人數</div></div>
                  <div className="stat-card">
                    <div className="value">{sortedM.length > 0 ? memberAvg[sortedM[0]] : '-'}</div>
                    <div className="label">最高平均分</div>
                  </div>
                  <div className="stat-card"><div className="value">{comments.length}</div><div className="label">留言數</div></div>
                </div>

                {/* Bar chart */}
                <div className="chart-card" style={{ marginBottom: 24 }}>
                  <h3>📊 {curSession.name} - 各成員平均被評分</h3>
                  <Bar data={sessionBar} options={{
                    indexAxis: 'y',
                    scales: { x: { min: 0, max: 5, ticks: { stepSize: 1 } } },
                    plugins: { legend: { display: false } }
                  }} />
                </div>

                {/* Score matrix */}
                <div className="chart-card" style={{ marginBottom: 24, overflowX: 'auto' }}>
                  <h3>📝 評分矩陣</h3>
                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid #e0e0de' }}>
                        <th style={{ padding: 6, textAlign: 'left', whiteSpace: 'nowrap' }}>評分者 ↓ / 被評者 →</th>
                        {members.map(m => <th key={m} style={{ padding: 6, textAlign: 'center', whiteSpace: 'nowrap' }}>{m}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {records.map((r, i) => (
                        <tr key={i} style={{ borderBottom: '1px solid #f0f0f0' }}>
                          <td style={{ padding: 6, fontWeight: 500, whiteSpace: 'nowrap' }}>{r.rater}</td>
                          {members.map(m => {
                            const isSelf = r.rater === m;
                            const score = r.scores[m];
                            return (
                              <td key={m} style={{
                                padding: 6, textAlign: 'center',
                                background: isSelf ? '#f5f5f4' : (score != null ? `rgba(194,213,48,${score / 7})` : ''),
                                color: isSelf ? '#aaa' : '#333',
                                fontWeight: score != null ? 500 : 400
                              }}>{isSelf ? '本人' : (score ?? '-')}</td>
                            );
                          })}
                        </tr>
                      ))}
                      {/* Average row */}
                      <tr style={{ borderTop: '2px solid #e0e0de', fontWeight: 600 }}>
                        <td style={{ padding: 6 }}>各人平均</td>
                        {members.map(m => (
                          <td key={m} style={{ padding: 6, textAlign: 'center', color: '#a8ba20' }}>{memberAvg[m]}</td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Individual score detail */}
                <div className="chart-card" style={{ marginBottom: 24, overflowX: 'auto' }}>
                  <h3>📋 個人得分明細</h3>
                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.9rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid #e0e0de', textAlign: 'left' }}>
                        <th style={{ padding: 8 }}>成員</th>
                        <th style={{ padding: 8 }}>平均分</th>
                        <th style={{ padding: 8 }}>各筆評分</th>
                        <th style={{ padding: 8 }}>評價</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedM.map(m => {
                        const scores = records.map(r => r.scores[m]).filter(s => s !== undefined);
                        return (
                          <tr key={m} style={{ borderBottom: '1px solid #f0f0f0' }}>
                            <td style={{ padding: 8, fontWeight: 500 }}>{m}</td>
                            <td style={{ padding: 8, fontWeight: 600, color: '#a8ba20' }}>{memberAvg[m]}</td>
                            <td style={{ padding: 8, color: '#6d6e71' }}>{scores.join(', ')}</td>
                            <td style={{ padding: 8, width: '25%' }}>
                              <div style={{ background: '#eef2d0', borderRadius: 4, overflow: 'hidden', height: 8 }}>
                                <div style={{ width: `${(memberAvg[m] / 5) * 100}%`, height: '100%', background: '#c2d530', borderRadius: 4 }} />
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {/* Comments */}
                {comments.length > 0 && (
                  <div className="chart-card" style={{ marginBottom: 24 }}>
                    <h3>💬 給組員的話</h3>
                    <div style={{ marginTop: 12 }}>
                      {comments.map((c, i) => (
                        <div key={i} style={{ padding: '10px 0', borderBottom: '1px solid #f0f0f0' }}>
                          <span style={{ color: '#6d6e71', fontSize: '0.85rem', marginRight: 8 }}>{c.rater}</span>
                          <span style={{ fontWeight: 500 }}>{c.comment}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </>
            );
          })()}
        </>
      )}
    </div>
  );
}
