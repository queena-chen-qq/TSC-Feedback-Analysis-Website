import React, { useState } from 'react';

const overlay = {
  position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
  background: 'rgba(0,0,0,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
};
const modal = {
  background: '#fff', borderRadius: 12, padding: 24, width: '90%', maxWidth: 480, maxHeight: '70vh', overflow: 'auto',
  boxShadow: '0 8px 32px rgba(0,0,0,0.2)'
};

export default function DeleteModal({ items, onDelete, onClose, title }) {
  const [selected, setSelected] = useState(new Set());
  const [deleting, setDeleting] = useState(false);

  const toggle = (item) => {
    const next = new Set(selected);
    next.has(item) ? next.delete(item) : next.add(item);
    setSelected(next);
  };

  const toggleAll = () => {
    setSelected(selected.size === items.length ? new Set() : new Set(items));
  };

  const handleDelete = async () => {
    if (selected.size === 0) return;
    setDeleting(true);
    await onDelete([...selected]);
    setDeleting(false);
    onClose();
  };

  return (
    <div style={overlay} onClick={onClose}>
      <div style={modal} onClick={e => e.stopPropagation()}>
        <h3 style={{ margin: '0 0 16px', color: '#3a3a3a' }}>{title || '選擇要刪除的資料'}</h3>

        <label style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0', borderBottom: '2px solid #e0e0de', cursor: 'pointer', fontWeight: 600, color: '#6d6e71' }}>
          <input type="checkbox" checked={selected.size === items.length && items.length > 0} onChange={toggleAll} style={{ width: 18, height: 18, accentColor: '#c2d530' }} />
          全選 ({items.length} 筆)
        </label>

        {items.map(item => (
          <label key={item} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 0', borderBottom: '1px solid #f0f0f0', cursor: 'pointer' }}>
            <input type="checkbox" checked={selected.has(item)} onChange={() => toggle(item)} style={{ width: 18, height: 18, accentColor: '#c2d530' }} />
            <span style={{ fontSize: '0.9rem', wordBreak: 'break-all' }}>{item}</span>
          </label>
        ))}

        <div style={{ display: 'flex', gap: 12, marginTop: 20, justifyContent: 'flex-end' }}>
          <button onClick={onClose} className="btn" style={{ background: '#e0e0de', color: '#6d6e71' }}>取消</button>
          <button onClick={handleDelete} className="btn" disabled={selected.size === 0 || deleting}
            style={{ background: selected.size > 0 ? '#c45040' : '#ccc', color: '#fff' }}>
            {deleting ? '刪除中...' : `刪除 (${selected.size})`}
          </button>
        </div>
      </div>
    </div>
  );
}
