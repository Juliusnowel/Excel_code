/************************************************************
 * kanban.gs — One-file Kanban Suite (ALL-IN-ONE)
 * - Mover v2 (gated, strict reconcile, force move)
 * - HTML Task Editor (save/edit/delete; NOTE-based, no column changes)
 * - Add New Task (content field -> HTML saved in NOTE)  ← (form & editor redone)
 * - Notifier (click-only Discord; uses Script Property fallback)
 * - Private Views (per-assignee Overview)
 * - Unified Menus
 *
 * All symbols use the KNB_ prefix to avoid clashes.
 ************************************************************/

/* =========================
   Small array-index helper (PVX)
========================= */
function KNB_indexFromArray_(headersArr){
  const m={}; (headersArr||[]).forEach((h,i)=>{ h=String(h||'').trim(); if(h) m[h]=i+1; });
  return m;
}