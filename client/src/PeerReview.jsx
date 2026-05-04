import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, PointElement, LineElement, Tooltip, Legend
} from 'chart.js';
import { Bar, Line } from 'react-chartjs-2';
import * as XLSX from 'xlsx';

ChartJS.register(CategoryScale, LinearScale, BarElement, PointElement, LineElement, Tooltip, Legend);

const PALETTE = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#4a4b4d','#d4c85a','#3a3a3a','#b8cc28','#e0a050','#c45040','#7cb342','#558b2f'];
const LINE_COLORS = ['#c2d530','#a8ba20','#6d6e71','#8faa1b','#e0a050','#c45040','#4a4b4d','#7cb342','#558b2f','#d4c85a','#b8cc28','#3a3a3a'];

function detectGroup(filename) {
  for (const g of ['大健康', '半導體', '綠能', '幕僚']) {
    if (filename.includes(g)) return g;
  }
  return '未分類';
}

function parseMultiSheetExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const rawName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');
        const group = detectGroup(rawName);

        // Helper: find the header row (the one containing '組員姓名' or '評分者')
        function findDataRows(sheet, keyword) {
          const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          let headerIdx = -1;
          for (let i = 0; i < Math.min(allRows.length, 10); i++) {
            const row = allRows[i];
            if (row && row.some(cell => String(cell ?? '').includes(keyword))) {
              headerIdx = i;
              break;
            }
          }
          if (headerIdx === -1) return { headers: [], rows: [] };
          const headers = allRows[headerIdx].map(h => String(h ?? '').trim());
          const rows = [];
          for (let i = headerIdx + 1; i < allRows.length; i++) {
            const r = allRows[i];
            if (!r || r.every(c => c === undefined || c === null || String(c).trim() === '')) continue;
            const obj = {};
            headers.forEach((h, ci) => { if (h) obj[h] = r[ci]; });
            rows.push(obj);
          }
          return { headers, rows };
        }

        // Parse overview sheet
        const overviewSheet = wb.Sheets[wb.SheetNames[0]];
        const { headers: ovHeaders, rows: ovRows } = findDataRows(overviewSheet, '組員姓名');
        const overview = [];
        if (ovRows.length > 0) {
          const nameCol = ovHeaders.find(h => h.includes('姓名')) || ovHeaders[0];
          const avgCol = ovHeaders.find(h => h.includes('總平均')) || '';
          const sessionCols = ovHeaders.filter(h => h && h !== nameCol && h !== avgCol && !h.includes('組'));

          ovRows.forEach(row => {
            const name = String(row[nameCol] ?? '').trim();
            if (!name) return;
            const sessions = {};
            sessionCols.forEach(col => {
              const val = parseFloat(row[col]);
              if (!isNaN(val)) sessions[col] = val;
            });
            const totalAvg = avgCol ? parseFloat(row[avgCol]) : 0;
            overview.push({ name, sessions, totalAvg: isNaN(totalAvg) ? 0 : totalAvg });
          });
        }

        // Parse each session sheet
        const sessions = [];
        for (let i = 1; i < wb.SheetNames.length; i++) {
          const sheetName = wb.SheetNames[i];
          const sheet = wb.Sheets[sheetName];
          const { headers, rows } = findDataRows(sheet, '評分者');
          if (rows.length === 0) continue;

          const raterCol = headers.find(h => h.includes('評分者')) || headers[0];
          const commentCol = headers.find(h => h.includes('留言') || h.includes('想對') || h.includes('說的話')) || '';
          const memberCols = headers.filter(h => h && h !== raterCol && h !== commentCol && !h.includes('平均'));

          const records = [];
          rows.forEach(row => {
            const rater = String(row[raterCol] ?? '').trim();
            if (!rater || rater.includes('平均')) return;

            const scores = {};
            memberCols.forEach(col => {
              const val = String(row[col] ?? '').trim();
              if (val === '本人') { /* skip self */ }
              else if (val && val !== '-' && !isNaN(Number(val))) scores[col] = Number(val);
            });

            const comment = commentCol ? String(row[commentCol] ?? '').trim() : '';
            records.push({
              rater,
              scores,
              comment: (comment && comment !== '無') ? comment : ''
            });
          });

          sessions.push({ name: sheetName, members: memberCols, records });
        }

        resolve({ group, overview, sessions, batch: rawName });
      } catch (err) { reject(err.message); }
    };
    reader.onerror = () => reject('檔案讀取失敗');
    reader.readAsArrayBuffer(file);
  });
}

export default function PeerReview() {
  const [allGroups, setAllGroups] = useState({}); // { groupName: { overview, sessions, batch } }
  const [selectedGroup, setSelectedGroup] = useState('');
  const [selectedSession, setSelectedSession] = useState('__overview__');
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');

  // Load from API
  const fetchAll = async () => {
    const res = await fetch('/api/peers');
    if (!res.ok) return;
    const items = await res.json();
    // Reconstruct groups from stored data
    const groups = {};
    items.forEach(item => {
      if (item.groupData) {
        groups[item.group] = item.groupData;
      }
    });
    setAllGroups(groups);
    const gNames = Object.keys(groups).sort();
    if (gNames.length > 0 && !selectedGroup) setSelectedGroup(gNames[0]);
  };

  useEffect(() => { fetchAll(); }, []);

  const handleUpload = async (e) => {
    e.preventDefault();
    const fi = e.target.querySelector('input[type="file"]');
    if (!fi.files[0]) return;
    setUploading(true); setMessage('');
    try {
      const result = await parseMultiSheetExcel(fi.files[0]);
      // Save as one record per group
      const record = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        batch: result.batch,
        group: result.group,
        groupData: { overview: result.overview, sessions: result.sessions, batch: result.batch }
      };
      await fetch('/api/peers', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ batch: result.group, records: [record] })
      });
      setMessage(`成功匯入 ${result.group} 互評資料`);
      setAllGroups(prev => ({ ...prev, [result.group]: record.groupData }));
      setSelectedGroup(result.group);
      setSelectedSession('__overview__');
    } catch (err) { setMessage(typeof err === 'string' ? err : '匯入失敗'); }
    setUploading(false); fi.value = '';
  };

  const handleClear = async () => {
    if (!confirm('確定要清除所有互評資料嗎？')) return;
    await fetch('/api/peers', { method: 'DELETE' });
    setAllGroups({}); setSelectedGroup(''); setSelectedSession('__overview__');
    setMessage('資料已清除');
  };

  const groupNames = Object.keys(allGroups).sort();
  const currentData = allGroups[selectedGroup];
  const overview = currentData?.overview || [];
  const sessions = currentData?.sessions || [];
  const currentSession = sessions.find(s => s.name === selectedSession);

  // Overview chart: line chart showing each member's score trend
  const sessionLabels = overview.length > 0 ? Object.keys(overview[0].sessions) : [];
  const lineChart = overview.length > 0 && sessionLabels.length > 1 ? {
    labels: sessionLabels.map(l => l.replace(/第\d+堂 \(/, '').replace(')', '')),
    datasets: overview.map((m, i) => ({
      label: m.name,
      data: sessionLabels.map(l => m.sessions[l] ?? null),
      borderColor: LINE_COLORS[i % LINE_COLORS.length],
      backgroundColor: LINE_COLORS[i % LINE_COLORS.length],
      tension: 0.3,
      pointRadius: 4,
      spanGaps: true
    }))
  } : null;

  // Overview bar chart: total average
  const sortedOverview = [...overview].sort((a, b) => b.totalAvg - a.totalAvg);
  const avgBarChart = sortedOverview.length > 0 ? {
    labels: sortedOverview.map(m => m.name),
    datasets: [{ label: '總平均', data: sortedOverview.map(m => m.totalAvg), backgroundColor: PALETTE.slice(0, sortedOverview.length) }]
  } : null;

  return (
    <div>
      <form className="upload-section" onSubmit={handleUpload}>
        <input type="file" accept=".xlsx,.xls,.csv" aria-label="選擇互評 Excel 檔案" />
        <button className="btn btn-primary" type="submit" disabled={uploading}>{uploading ? '上傳中...' : '匯入互評 Excel'}</button>
        <button className="btn btn-danger" type="button" onClick={handleClear}>清除資料</button>
        {message && <span style={{ color: '#a8ba20', fontWeight: 500 }}>{message}</span>}
      </form>
      <p style={{ color: '#999', fontSize: '0.85rem', marginBottom: 16 }}>檔名需包含組別名稱（大健康、半導體、綠能、幕僚）</p>

      {groupNames.length > 0 && (
        <>
          {/* Group tabs */}
          <div style={{ display: 'flex', gap: 0, marginBottom: 20, borderBottom: '2px solid #e0e0de' }}>
            {groupNames.map(g => (
              <button key={g} onClick={() => { setSelectedGroup(g); setSelectedSession('__overview__'); }} style={{
                padding: '10px 24px', border: 'none', cursor: 'pointer', fontSize: '0.95rem', fontWeight: 500,
                background: selectedGroup === g ? '#fff' : 'transparent',
                color: selectedGroup === g ? '#a8ba20' : '#6d6e71',
                borderBottom: selectedGroup === g ? '3px solid #c2d530' : '3px solid transparent',
                marginBottom: -2, borderRadius: '8px 8px 0 0'
              }}>{g}</button>
            ))}
          </div>

          {/* Session tabs */}
          {currentData && (
            <div className="controls" style={{ flexWrap: 'wrap', gap: 6 }}>
              <button onClick={() => setSelectedSession('__overview__')} className="btn" style={{
                padding: '6px 14px', fontSize: '0.85rem',
                background: selectedSession === '__overview__' ? '#c2d530' : '#e0e0de',
                color: selectedSession === '__overview__' ? '#fff' : '#6d6e71'
              }}>📊 總覽</button>
              {sessions.map(s => (
                <button key={s.name} onClick={() => setSelectedSession(s.name)} className="btn" style={{
                  padding: '6px 14px', fontSize: '0.85rem',
                  background: selectedSession === s.name ? '#c2d530' : '#e0e0de',
                  color: selectedSession === s.name ? '#fff' : '#6d6e71'
                }}>{s.name}</button>
              ))}
            </div>
          )}
        </>
      )}

      {/* OVERVIEW VIEW */}
      {currentData && selectedSession === '__overview__' && (
        <>
          {/* Total average bar chart */}
          {avgBarChart && (
            <div className="chart-card" style={{ marginBottom: 24 }}>
              <h3>📊 各成員總平均分</h3>
              <Bar data={avgBarChart} options={{ indexAxis: 'y', scales: { x: { min: 0, max: 5, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } }} />
            </div>
          )}

          {/* Line chart: score trend */}
          {lineChart && (
            <div className="chart-card" style={{ marginBottom: 24 }}>
              <h3>📈 各成員歷次分數趨勢</h3>
              <Line data={lineChart} options={{
                scales: { y: { min: 0, max: 5, ticks: { stepSize: 1 } } },
                plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 8, font: { size: 11 } } } }
              }} />
            </div>
          )}

          {/* Overview table */}
          <div className="chart-card" style={{ overflowX: 'auto', marginBottom: 24 }}>
            <h3>📋 各堂課平均分數總覽</h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.9rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #eee', textAlign: 'left' }}>
                  <th style={{ padding: 8, whiteSpace: 'nowrap' }}>組員姓名</th>
                  {sessionLabels.map(l => <th key={l} style={{ padding: 8, whiteSpace: 'nowrap', textAlign: 'center' }}>{l}</th>)}
                  <th style={{ padding: 8, whiteSpace: 'nowrap', textAlign: 'center', color: '#a8ba20', fontWeight: 700 }}>總平均</th>
                </tr>
              </thead>
              <tbody>
                {sortedOverview.map(m => (
                  <tr key={m.name} style={{ borderBottom: '1px solid #f0f0f0' }}>
                    <td style={{ padding: 8, fontWeight: 500 }}>{m.name}</td>
                    {sessionLabels.map(l => {
                      const v = m.sessions[l];
                      return <td key={l} style={{ padding: 8, textAlign: 'center', color: v >= 4.5 ? '#a8ba20' : v < 3.5 ? '#c45040' : '#333' }}>{v?.toFixed(2) ?? '-'}</td>;
                    })}
                    <td style={{ padding: 8, textAlign: 'center', fontWeight: 700, color: '#a8ba20' }}>{m.totalAvg.toFixed(2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {/* SESSION DETAIL VIEW */}
      {currentSession && selectedSession !== '__overview__' && (() => {
        const members = currentSession.members;
        const records = currentSession.records;
        const comments = records.filter(r => r.comment);

        // Compute averages per member
        const memberAvg = {};
        members.forEach(m => {
          const scores = records.map(r => r.scores[m]).filter(s => s !== undefined);
          memberAvg[m] = scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
        });
        const sortedMembers = [...members].sort((a, b) => memberAvg[b] - memberAvg[a]);

        const barChart = {
          labels: sortedMembers,
          datasets: [{ label: '平均被評分', data: sortedMembers.map(m => memberAvg[m]), backgroundColor: PALETTE.slice(0, sortedMembers.length) }]
        };

        return (
          <>
            <div className="chart-card" style={{ marginBottom: 24 }}>
              <h3>📊 {currentSession.name} - 各成員平均被評分</h3>
              <Bar data={barChart} options={{ indexAxis: 'y', scales: { x: { min: 0, max: 5, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } }} />
            </div>

            <div className="chart-card" style={{ marginBottom: 24, overflowX: 'auto' }}>
              <h3>📝 評分矩陣</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12, fontSize: '0.85rem' }}>
                <thead>
                  <tr style={{ borderBottom: '2px solid #eee' }}>
                    <th style={{ padding: 6, textAlign: 'left' }}>評分者 ↓ / 被評者 →</th>
                    {members.map(m => <th key={m} style={{ padding: 6, whiteSpace: 'nowrap', textAlign: 'center' }}>{m}</th>)}
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
                            background: isSelf ? '#f5f5f4' : (score ? `rgba(194,213,48,${score/7})` : ''),
                            color: isSelf ? '#999' : '#333', fontWeight: score ? 500 : 400
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

      {groupNames.length === 0 && (
        <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
          <p style={{ fontSize: '1.2rem' }}>尚無互評資料，請先匯入 Excel 檔案</p>
          <p style={{ marginTop: 8, fontSize: '0.9rem' }}>每組一份 Excel，包含總覽和各堂課互評 sheet</p>
        </div>
      )}
    </div>
  );
}
