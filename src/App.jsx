import React, { useReducer, useMemo, useCallback, useRef, useState, useEffect } from 'react';
import Papa from 'papaparse';
import { neon } from '@neondatabase/serverless';
import PptxGenJS from 'pptxgenjs';
import {
  LineChart, Line, AreaChart, Area, BarChart, Bar,
  ComposedChart, PieChart, Pie, Cell, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine
} from 'recharts';

// ─── RESPONSIVE STYLES ───
const ResponsiveStyles = () => (
  <style>{`
    .crm-tab-nav::-webkit-scrollbar { display: none; }
    .crm-kpi-grid { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
    .crm-2col-grid { grid-template-columns: 1fr 1fr; }
    .crm-heatmap-grid { grid-template-columns: 2fr 1fr; }
    @media (max-width: 768px) {
      .crm-header { flex-direction: column; gap: 10px; padding: 12px 16px !important; }
      .crm-header-left { gap: 8px !important; }
      .crm-header-left h1 { font-size: 17px !important; }
      .crm-header-left p { font-size: 11px !important; }
      .crm-header-right { width: 100%; justify-content: space-between !important; flex-wrap: wrap; }
      .crm-header-right .crm-pres-btn { display: none; }
      .crm-main { padding: 0 12px 24px !important; }
      .crm-kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; gap: 8px !important; }
      .crm-2col-grid { grid-template-columns: 1fr !important; }
      .crm-card { padding: 12px !important; }
      .crm-card h3 { font-size: 13px !important; }
      .crm-dp-panel { flex-direction: column; min-width: unset !important; width: calc(100vw - 24px) !important; max-height: 80vh; overflow-y: auto; left: -12px !important; }
      .crm-dp-presets { width: 100% !important; border-right: none !important; border-bottom: 1px solid #F0ECE3; flex-direction: row !important; flex-wrap: wrap; padding: 8px !important; gap: 4px !important; }
      .crm-dp-presets > div:first-child { display: none; }
      .crm-dp-presets button { padding: 4px 10px !important; font-size: 11px !important; border-left: none !important; border-radius: 4px !important; background: #F6EDDA !important; }
      .crm-dp-presets button[style*="font-weight: 600"] { background: #124A2B11 !important; }
      .crm-dp-calendar { padding: 12px !important; }
      .crm-dp-months { gap: 12px !important; flex-direction: column !important; }
      .crm-chart-card { padding: 12px !important; }
      .crm-chart-card .recharts-wrapper { font-size: 10px; }
      .crm-table-wrap { font-size: 11px !important; }
      .crm-tab-nav { gap: 2px !important; padding: 6px 0 !important; }
      .crm-tab-nav button { padding: 6px 10px !important; font-size: 11px !important; }
      .crm-pitstop-kanban { grid-template-columns: repeat(2, 1fr) !important; }
      .crm-pres-editor { flex-direction: column !important; height: 95vh !important; }
      .crm-pres-sidebar { width: 100% !important; max-height: 180px !important; border-right: none !important; border-bottom: 1px solid #F0ECE3 !important; }
      .crm-pres-editor-panel { padding: 12px !important; }
      .crm-calendar-grid > div { min-height: 60px !important; }
      .crm-calendar-grid .crm-cal-item { font-size: 9px !important; }
    }
    @media (max-width: 480px) {
      .crm-header-left .crm-logo { height: 22px !important; }
      .crm-pitstop-kanban { grid-template-columns: 1fr !important; }
      .crm-calendar-grid > div { min-height: 50px !important; }
      .crm-calendar-grid .crm-cal-item { display: none !important; }
      .crm-calendar-grid .crm-cal-count { display: block !important; }
    }
  `}</style>
);

// ─── COLOR TOKENS (omni.pet brand) ───
const C = {
  primary: '#124A2B',
  success: '#18917B',
  danger: '#D81F26',
  warning: '#F59E0B',
  secondary: '#18917B',
  info: '#18917B',
  textPrimary: '#272D45',
  textSecondary: '#676986',
  textTertiary: '#9CA3AF',
  cardBg: '#FFFFFF',
  pageBg: '#F6EDDA',
  cardBorder: '#E5E5EB',
  divider: '#F0ECE3',
};

const CRM_CHANNEL_COLORS = {
  'Email Campaign': '#124A2B',
  'Welcome Series': '#18917B',
  'Win-Back 60d': '#F59E0B',
  'Abandoned Cart': '#D81F26',
  'Post-Purchase Upsell': '#2D8B6E',
  'Re-Engagement 90d': '#676986',
  'WhatsApp': '#25D366',
  'SMS Blast': '#18917B',
  'Personal Call': '#E67E22',
  'Gift-with-Purchase': '#C0392B',
  'Surprise & Delight': '#2D8B6E',
};

const LIFECYCLE_COLORS = {
  'New': '#18917B',
  'Active': '#2D8B6E',
  'At-Risk': '#F59E0B',
  'Lapsed': '#D81F26',
};

const TIER_COLORS = {
  '2nd Order': '#3B82F6',
  '3rd Order': '#F59E0B',
  '6th Order': '#10B981',
};

const CHANNEL_DEFS = [
  { key: 'Emails', color: '#124A2B' },
  { key: 'SMS', color: '#3B82F6' },
  { key: 'WhatsApp', color: '#25D366' },
  { key: 'Postcards', color: '#F59E0B' },
];

const MILESTONE_TIERS = [
  { key: '2ndOrder', label: '2nd Order' },
  { key: '3rdOrder', label: '3rd Order' },
  { key: '6thOrder', label: '6th Order' },
];

const OmniPetLogo = ({ height = 28 }) => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 45" style={{ height, width: 'auto' }}>
    <defs><clipPath id="omni-clip"><path d="M.5 1H200v42.8H.5z"/></clipPath></defs>
    <g style={{ clipPath: 'url(#omni-clip)' }}>
      <path fill="#d81f26" d="M.5 22.1C.5 9.8 11 1 28.4 1s27.8 8.7 27.8 21.1-10.5 21.7-27.8 21.7S.5 34 .5 22.1m36.1.1c0-4.7-3.3-7.8-8.3-7.8s-8.3 3-8.3 7.8 3.3 8.1 8.3 8.1 8.3-3.3 8.3-8.1"/>
      <path fill="#d81f26" d="M128.4 2.4h22.4l11.1 22.3V2.4h15.4v40h-22l-11.4-22.5v22.5h-15.4v-40z"/>
      <path fill="#d81f26" d="M181.9 2.3H200v40.1h-18.1z"/>
      <path fill="#d81f26" d="M106.2 16.1h.5c5.7 0 6.9 4.4 7.1 7.3v19.1H125l-2.7-40H96.5l-5.1 19.2-4.9-19.2H60.4l-2.2 40h11.2V24.4c0-2.6.7-8.3 7.1-8.3s7.1 5.7 7.1 8.3V31c5.2-1.4 10.7-1.4 16 0v-6.7c0-2.5.6-7.9 6.5-8.2"/>
      <path fill="#d81f26" d="M100.4 35.9c-1.8 0-3.3 1.4-3.3 3.1s1.4 3.3 3.1 3.3h.2c1.8 0 3.3-1.3 3.4-3.1 0-1.8-1.3-3.3-3.1-3.4h-.3"/>
      <path fill="#d81f26" d="M82.6 36c-1.8 0-3.2 1.5-3.1 3.3s1.5 3.2 3.3 3.1c1.7 0 3.1-1.5 3.1-3.2 0-1.8-1.5-3.2-3.3-3.2"/>
    </g>
  </svg>
);

// ─── DATA DEFAULTS (empty — import your own via AI Importer or CSV) ───
const DEMO_REVENUE = [];
const DEMO_SUBSCRIPTIONS = [];
const DEMO_EMAIL_FLOWS = [];
const DEMO_LOYALTY = [];
const DEMO_MILESTONE_PRODUCTS = [];
const DEMO_SEGMENTS = [];
const DEMO_OUTREACH = [];
const DEMO_WHATSAPP_FLOWS = [];
const DEMO_POSTCARD_FLOWS = [];
const DEMO_CHANNEL_COSTS = [];
const DEMO_PRODUCT_CHURN = [];
const DEMO_BEFORE_AFTER = [];
const DEMO_HOLDOUT_TESTS = [];
const DEMO_ACTIVITY_ROI = [];


// ─── UTILITIES ───
const formatCurrency = (v) => v == null ? '—' : new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP', maximumFractionDigits: 0 }).format(v);
const formatCurrencyDecimal = (v) => v == null ? '—' : new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
const formatPercent = (v) => v == null ? '—' : `${v.toFixed(1)}%`;
const formatMultiplier = (v) => v == null ? '—' : `${v.toFixed(1)}x`;
const formatNumber = (v) => v == null ? '—' : new Intl.NumberFormat('en-GB').format(v);
const calcDelta = (arr, key) => {
  if (!arr || arr.length < 2) return null;
  const curr = arr[arr.length - 1][key];
  const prev = arr[arr.length - 2][key];
  if (prev === 0 || prev == null || curr == null) return null;
  return ((curr - prev) / Math.abs(prev)) * 100;
};
const pctChange = (curr, prev) => {
  if (prev == null || curr == null || prev === 0) return null;
  return ((curr - prev) / Math.abs(prev)) * 100;
};
const COMP_LABELS = { previous_period: 'vs prev period', previous_month: 'vs prev month', previous_year: 'vs prev year' };
const lastN = (arr, n) => arr?.slice(-n) ?? [];

function getChannelCostAndROI(channelCosts, channelName, totalRevenue, start, end) {
  const startMonth = start?.slice(0, 7);
  const endMonth = end?.slice(0, 7);
  const relevantCosts = (channelCosts || []).filter(c => {
    if (c.channel !== channelName) return false;
    if (startMonth && c.month < startMonth) return false;
    if (endMonth && c.month > endMonth) return false;
    return true;
  });
  const totalCost = relevantCosts.reduce((sum, c) => sum + (Number(c.cost) || 0), 0);
  const roi = totalCost > 0 ? totalRevenue / totalCost : 0;
  return { totalCost, roi };
}

// ─── TABS ───
// ─── DATE RANGE UTILITIES ───
function filterByDateRange(data, start, end, dateField = 'week') {
  if (!data || !data.length) return data;
  return data.filter(row => {
    const val = row[dateField];
    if (!val) return false;
    // Monthly data: "2025-03" format — include if any part overlaps
    if (val.length === 7) {
      const monthStart = val + '-01';
      const d = new Date(val + '-01');
      const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0).toISOString().slice(0, 10);
      return monthEnd >= start && monthStart <= end;
    }
    return val >= start && val <= end;
  });
}

function computeComparisonRange(start, end, comparison) {
  if (comparison === 'none') return null;
  const s = new Date(start), e = new Date(end);
  const days = Math.round((e - s) / 86400000);
  if (comparison === 'previous_period') {
    const cs = new Date(s); cs.setDate(cs.getDate() - days - 1);
    const ce = new Date(s); ce.setDate(ce.getDate() - 1);
    return { start: cs.toISOString().slice(0, 10), end: ce.toISOString().slice(0, 10) };
  }
  if (comparison === 'previous_month') {
    const cs = new Date(s); cs.setMonth(cs.getMonth() - 1);
    const ce = new Date(e); ce.setMonth(ce.getMonth() - 1);
    return { start: cs.toISOString().slice(0, 10), end: ce.toISOString().slice(0, 10) };
  }
  if (comparison === 'previous_year') {
    const cs = new Date(s); cs.setFullYear(cs.getFullYear() - 1);
    const ce = new Date(e); ce.setFullYear(ce.getFullYear() - 1);
    return { start: cs.toISOString().slice(0, 10), end: ce.toISOString().slice(0, 10) };
  }
  return null;
}

function computeMetricGroupValues(state, start, end) {
  const r = {};
  // Email & Flows (weekly)
  const fE = filterByDateRange(state.emailFlows, start, end, 'week');
  const camps = fE.filter(x => x.type === 'Campaign');
  r.emailFlows = {
    totalEmailRevenue: fE.reduce((s, x) => s + (x.revenue || 0), 0),
    flowRevenue: fE.filter(x => x.type === 'Flow').reduce((s, x) => s + (x.revenue || 0), 0),
    campaignRevenue: camps.reduce((s, x) => s + (x.revenue || 0), 0),
    avgOpenRate: camps.length ? camps.reduce((s, x) => s + (x.openRate || 0), 0) / camps.length : 0,
    avgCTR: camps.length ? camps.reduce((s, x) => s + (x.ctr || 0), 0) / camps.length : 0,
    avgUnsubRate: camps.length ? camps.reduce((s, x) => s + (x.unsubRate || 0), 0) / camps.length : 0,
    listSize: [...fE].reverse().find(x => x.listSize)?.listSize || 0,
  };
  // Loyalty (monthly)
  const fL = filterByDateRange(state.loyalty, start, end, 'month');
  const lL = fL.length ? fL[fL.length - 1] : null;
  r.loyalty = {
    totalMembers: lL?.totalMembers || 0, newEnrollments: lL?.newEnrollments || 0,
    redemptionRate: lL?.redemptionRate || 0, memberAOV: lL?.memberAOV || 0,
    nonMemberAOV: lL?.nonMemberAOV || 0,
    aovLift: lL && lL.nonMemberAOV > 0 ? ((lL.memberAOV - lL.nonMemberAOV) / lL.nonMemberAOV) * 100 : 0,
    memberRetentionRate: lL?.memberRetentionRate || 0, nonMemberRetentionRate: lL?.nonMemberRetentionRate || 0,
    tier6thOrderLTV: lL?.tier6thOrderLTV || 0,
    ltvLift: lL && lL.nonMemberLTV > 0 ? ((lL.tier6thOrderLTV - lL.nonMemberLTV) / lL.nonMemberLTV) * 100 : 0,
  };
  // Segments (monthly)
  const fS = filterByDateRange(state.segments, start, end, 'month');
  const lS = fS.length ? fS[fS.length - 1] : null;
  r.segments = {
    totalCustomers: lS?.totalCustomers || 0, segActive: lS?.segActive || 0,
    segAtRisk: lS?.segAtRisk || 0, segLapsed: lS?.segLapsed || 0, segNew: lS?.segNew || 0,
    avgRFMScore: lS?.avgRFMScore || 0, migratedAtRiskToActive: lS?.migratedAtRiskToActive || 0,
  };
  // Revenue (weekly)
  const fR = filterByDateRange(state.revenue, start, end, 'week');
  const lR = fR.length ? fR[fR.length - 1] : null;
  r.revenue = {
    totalRevenue: lR?.totalRevenue || 0, netRevenue: lR?.netRevenue || 0,
    subscriptionRevenue: lR?.subscriptionRevenue || 0, oneTimeRevenue: lR?.oneTimeRevenue || 0,
    totalOrders: lR?.totalOrders || 0, aov: lR?.aov || 0,
  };
  // Subscriptions (monthly)
  const fSub = filterByDateRange(state.subscriptions, start, end, 'month');
  const lSub = fSub.length ? fSub[fSub.length - 1] : null;
  r.subscriptions = {
    activeSubscribers: lSub?.activeSubscribers || 0, mrr: lSub?.mrr || 0,
    churnRate: lSub?.churnRate || 0, newSubscribers: lSub?.newSubscribers || 0,
    churnedSubscribers: lSub?.churnedSubscribers || 0, ltv: lSub?.ltv || 0,
  };
  // Outreach (weekly)
  const fO = filterByDateRange(state.outreach, start, end, 'week');
  const outRev = fO.reduce((s, x) => s + (x.revenue || 0), 0);
  const outCost = fO.reduce((s, x) => s + (x.cost || 0), 0);
  const latOW = [...new Set(fO.map(x => x.week))].sort().pop();
  const latOD = fO.filter(x => x.week === latOW);
  r.outreach = {
    outreachRevenue: outRev, outreachCost: outCost,
    outreachROAS: outCost > 0 ? outRev / outCost : 0,
    waResponseRate: latOD.find(x => x.channel === 'WhatsApp')?.responseRate || 0,
    smsConvRate: latOD.find(x => x.channel === 'SMS Blast')?.conversionRate || 0,
  };
  // Incrementality (not date-filtered)
  const totIncRev = state.activityROI.reduce((s, x) => s + x.incrementalRevenue, 0);
  const totCost = state.activityROI.reduce((s, x) => s + x.totalCost, 0);
  const topLift = state.beforeAfter.reduce((m, x) => x.lift > m.lift ? x : m, { lift: 0, activity: '' });
  r.incrementality = {
    totalIncrementalRevenue: totIncRev,
    avgIncrementalROI: totCost > 0 ? totIncRev / totCost : 0,
    activeHoldoutTests: state.holdoutTests.filter(t => t.status === 'active').length,
    highestLiftActivity: topLift.activity ? `${topLift.activity} +${topLift.lift.toFixed(0)}%` : 'None',
  };
  return r;
}

// Format date as YYYY-MM-DD in local timezone (avoids UTC shift from toISOString)
const fmtDate = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

function getDatePresets() {
  const now = new Date();
  const fmt = fmtDate;
  const sub = (d, n) => { const r = new Date(d); r.setDate(r.getDate() - n); return r; };
  const y = now.getFullYear(), m = now.getMonth(); // 0-indexed month
  const monthStart = new Date(y, m, 1);
  const monthEnd = new Date(y, m + 1, 0); // last day of current month
  const lastMonthStart = new Date(y, m - 1, 1);
  const lastMonthEnd = new Date(y, m, 0); // last day of previous month
  const quarterStart = new Date(y, Math.floor(m / 3) * 3, 1);
  const quarterEnd = new Date(y, Math.floor(m / 3) * 3 + 3, 0); // last day of current quarter
  return [
    { label: 'Last 7 days', start: fmt(sub(now, 6)), end: fmt(now) },
    { label: 'Last 14 days', start: fmt(sub(now, 13)), end: fmt(now) },
    { label: 'Last 30 days', start: fmt(sub(now, 29)), end: fmt(now) },
    { label: 'Last 90 days', start: fmt(sub(now, 89)), end: fmt(now) },
    { label: 'This month', start: fmt(monthStart), end: fmt(monthEnd) },
    { label: 'Last month', start: fmt(lastMonthStart), end: fmt(lastMonthEnd) },
    { label: 'This quarter', start: fmt(quarterStart), end: fmt(quarterEnd) },
    { label: 'Last 12 months', start: fmt(sub(now, 364)), end: fmt(now) },
    { label: 'All time', start: '2024-01-01', end: fmt(now) },
  ];
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
}

const TABS = [
  { key: 'overview', label: 'CRM Overview' },
  { key: 'email', label: 'Email & Flows' },
  { key: 'whatsapp', label: 'WhatsApp' },
  { key: 'postcard', label: 'Postcards' },
  { key: 'loyalty', label: 'Milestone Reward' },
  { key: 'segments', label: 'Segments & Lifecycle' },
  { key: 'incrementality', label: 'Incrementality' },
  { key: 'initiatives', label: 'Initiatives' },
  { key: 'import', label: 'Data Import' },
];

// ─── TIME PERIOD AGGREGATION ───
function aggregateByMonth(data, dateField = 'week') {
  const groups = {};
  data.forEach(row => {
    const d = row[dateField];
    if (!d) return;
    const month = d.slice(0, 7);
    if (!groups[month]) groups[month] = [];
    groups[month].push(row);
  });
  return groups;
}

function aggregateEmailFlowsByMonth(data) {
  const groups = aggregateByMonth(data, 'week');
  const result = [];
  Object.entries(groups).sort(([a], [b]) => a.localeCompare(b)).forEach(([month, rows]) => {
    const types = {};
    rows.forEach(r => {
      const key = r.type === 'Campaign' ? 'Campaign' : r.flowName;
      if (!types[key]) types[key] = { ...r, sends: 0, delivered: 0, opens: 0, clicks: 0, unsubscribes: 0, revenue: 0, conversions: 0 };
      types[key].sends += r.sends || 0;
      types[key].delivered += r.delivered || 0;
      types[key].opens += r.opens || 0;
      types[key].clicks += r.clicks || 0;
      types[key].unsubscribes += r.unsubscribes || 0;
      types[key].revenue += r.revenue || 0;
      types[key].conversions += r.conversions || 0;
    });
    Object.values(types).forEach(t => {
      t.week = month;
      t.openRate = t.sends > 0 ? (t.opens / t.sends) * 100 : 0;
      t.ctr = t.sends > 0 ? (t.clicks / t.sends) * 100 : 0;
      t.unsubRate = t.sends > 0 ? (t.unsubscribes / t.sends) * 100 : 0;
      if (t.type === 'Campaign') {
        const latestCamp = rows.filter(r => r.type === 'Campaign').pop();
        t.listSize = latestCamp?.listSize;
      }
      result.push(t);
    });
  });
  return result;
}

function aggregateOutreachByMonth(data) {
  const groups = aggregateByMonth(data, 'week');
  const result = [];
  Object.entries(groups).sort(([a], [b]) => a.localeCompare(b)).forEach(([month, rows]) => {
    const channels = {};
    rows.forEach(r => {
      if (!channels[r.channel]) channels[r.channel] = { ...r, sends: 0, delivered: 0, responses: 0, conversions: 0, revenue: 0, cost: 0 };
      channels[r.channel].sends += r.sends || 0;
      channels[r.channel].delivered += r.delivered || 0;
      channels[r.channel].responses += r.responses || 0;
      channels[r.channel].conversions += r.conversions || 0;
      channels[r.channel].revenue += r.revenue || 0;
      channels[r.channel].cost += r.cost || 0;
    });
    Object.values(channels).forEach(c => {
      c.week = month;
      c.responseRate = c.sends > 0 ? (c.responses / c.sends) * 100 : 0;
      c.conversionRate = c.sends > 0 ? (c.conversions / c.sends) * 100 : 0;
      result.push(c);
    });
  });
  return result;
}

function aggregateRevenueByMonth(data) {
  const groups = aggregateByMonth(data, 'week');
  return Object.entries(groups).sort(([a], [b]) => a.localeCompare(b)).map(([month, rows]) => ({
    week: month,
    totalRevenue: rows.reduce((s, r) => s + (r.totalRevenue || 0), 0),
    subscriptionRevenue: rows.reduce((s, r) => s + (r.subscriptionRevenue || 0), 0),
    oneTimeRevenue: rows.reduce((s, r) => s + (r.oneTimeRevenue || 0), 0),
    refunds: rows.reduce((s, r) => s + (r.refunds || 0), 0),
    netRevenue: rows.reduce((s, r) => s + (r.netRevenue || 0), 0),
    totalOrders: rows.reduce((s, r) => s + (r.totalOrders || 0), 0),
    aov: rows.reduce((s, r) => s + (r.totalRevenue || 0), 0) / Math.max(1, rows.reduce((s, r) => s + (r.totalOrders || 0), 0)),
  }));
}

// ─── STATE ───
// Keys that hold imported dashboard data and should persist in Neon
const DATA_KEYS = ['emailFlows','loyalty','segments','outreach','beforeAfter','holdoutTests','activityROI','revenue','subscriptions','milestoneProducts','whatsappFlows','postcardFlows','channelCosts','productChurn'];

const initialState = {
  activeTab: 'overview',
  emailFlows: DEMO_EMAIL_FLOWS,
  loyalty: DEMO_LOYALTY,
  segments: DEMO_SEGMENTS,
  outreach: DEMO_OUTREACH,
  beforeAfter: DEMO_BEFORE_AFTER,
  holdoutTests: DEMO_HOLDOUT_TESTS,
  activityROI: DEMO_ACTIVITY_ROI,
  revenue: DEMO_REVENUE,
  subscriptions: DEMO_SUBSCRIPTIONS,
  milestoneProducts: DEMO_MILESTONE_PRODUCTS,
  whatsappFlows: DEMO_WHATSAPP_FLOWS,
  postcardFlows: DEMO_POSTCARD_FLOWS,
  channelCosts: DEMO_CHANNEL_COSTS,
  productChurn: DEMO_PRODUCT_CHURN,
  lastUpdated: { emailFlows: null, loyalty: null, segments: null, outreach: null, beforeAfter: null, holdoutTests: null, activityROI: null, revenue: null, subscriptions: null, milestoneProducts: null, whatsappFlows: null, postcardFlows: null, channelCosts: null, productChurn: null },
  settingsOpen: false,
  tabPeriods: { overview: 'weekly', email: 'weekly', loyalty: 'monthly', segments: 'monthly', outreach: 'weekly', incrementality: 'all' },
  dateRange: { start: fmtDate(new Date(new Date().getFullYear(), new Date().getMonth(), 1)), end: fmtDate(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0)) },
  pendingDateRange: { start: fmtDate(new Date(new Date().getFullYear(), new Date().getMonth(), 1)), end: fmtDate(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0)) },
  comparison: 'none',
  pendingComparison: 'none',
  dateMode: 'fixed',
  pendingDateMode: 'fixed',
  datePickerOpen: false,
  segmentLinks: [],
  activityLog: [],
};

function reducer(state, action) {
  switch (action.type) {
    case 'SET_TAB': return { ...state, activeTab: action.payload };
    case 'LOAD_DATA': return { ...state, [action.source]: action.payload, lastUpdated: { ...state.lastUpdated, [action.source]: new Date().toLocaleString('en-GB') } };
    case 'APPEND_DATA': return { ...state, [action.source]: [...(state[action.source] || []), ...action.payload], lastUpdated: { ...state.lastUpdated, [action.source]: new Date().toLocaleString('en-GB') } };
    case 'RESET_DEMO': return { ...initialState, activeTab: state.activeTab, settingsOpen: state.settingsOpen, tabPeriods: state.tabPeriods };
    case 'CLEAR_ALL': return { ...state, emailFlows: [], loyalty: [], segments: [], outreach: [], beforeAfter: [], holdoutTests: [], activityROI: [], revenue: [], subscriptions: [], milestoneProducts: [], whatsappFlows: [], postcardFlows: [], channelCosts: [], productChurn: [], lastUpdated: Object.fromEntries(Object.keys(state.lastUpdated).map(k => [k, 'Cleared'])) };
    case 'TOGGLE_SETTINGS': return { ...state, settingsOpen: !state.settingsOpen };
    case 'SET_TAB_PERIOD': return { ...state, tabPeriods: { ...state.tabPeriods, [action.tab]: action.period } };
    case 'TOGGLE_DATE_PICKER': return { ...state, datePickerOpen: !state.datePickerOpen, pendingDateRange: state.dateRange, pendingComparison: state.comparison, pendingDateMode: state.dateMode };
    case 'SET_PENDING_DATE_RANGE': return { ...state, pendingDateRange: action.payload };
    case 'SET_PENDING_COMPARISON': return { ...state, pendingComparison: action.payload };
    case 'SET_PENDING_DATE_MODE': return { ...state, pendingDateMode: action.payload };
    case 'APPLY_DATE_RANGE': return { ...state, dateRange: state.pendingDateRange, comparison: state.pendingComparison, dateMode: state.pendingDateMode, datePickerOpen: false };
    case 'SET_DATE_RANGE': return { ...state, dateRange: action.payload, pendingDateRange: action.payload };
    case 'SET_COMPARISON': return { ...state, comparison: action.payload, pendingComparison: action.payload };
    case 'CANCEL_DATE_RANGE': return { ...state, pendingDateRange: state.dateRange, pendingComparison: state.comparison, pendingDateMode: state.dateMode, datePickerOpen: false };
    case 'SET_SEGMENT_LINKS': return { ...state, segmentLinks: action.payload };
    case 'ADD_SEGMENT_LINK': return { ...state, segmentLinks: [...state.segmentLinks, action.payload] };
    case 'UPDATE_SEGMENT_LINK': return { ...state, segmentLinks: state.segmentLinks.map(s => s.id === action.payload.id ? { ...s, ...action.payload } : s) };
    case 'DELETE_SEGMENT_LINK': return { ...state, segmentLinks: state.segmentLinks.filter(s => s.id !== action.payload) };
    case 'LOG_ACTIVITY': return { ...state, activityLog: [{ id: Date.now(), ...action.payload, timestamp: new Date().toISOString() }, ...state.activityLog].slice(0, 200) };
    case 'SET_ACTIVITY_LOG': return { ...state, activityLog: action.payload };
    default: return state;
  }
}

// ─── CSV TEMPLATES ───
const CSV_TEMPLATES = {
  emailFlows: { headers: ['week','type','flowName','sends','delivered','opens','openRate','clicks','ctr','unsubscribes','unsubRate','revenue','conversions','listSize'], sample: [['2025-03-24','Campaign','','47000','45120','19853','44.0','2888','6.40','52','0.12','11200','285','53400']] },
  loyalty: { headers: ['month','totalMembers','newEnrollments','pointsIssued','pointsRedeemed','redemptionRate','rewardsRedeemed','rewardCostGBP','revenueFromMembers','revenueFromNonMembers','memberAOV','nonMemberAOV','memberRetentionRate','nonMemberRetentionRate','tier2ndOrderMembers','tier2ndOrderAOV','tier2ndOrderLTV','tier3rdOrderMembers','tier3rdOrderAOV','tier3rdOrderLTV','tier6thOrderMembers','tier6thOrderAOV','tier6thOrderLTV','nonMemberLTV'], sample: [['2025-03','3680','580','736000','147200','20.0','442','13260','162400','48200','89.50','72.50','92.4','69.8','2400','85.20','170.40','880','93.80','281.40','400','108.60','651.60','116.00']] },
  segments: { headers: ['month','segNew','segActive','segAtRisk','segLapsed','totalCustomers','avgRFMScore','segNewRevenue','segActiveRevenue','segAtRiskRevenue','segLapsedRevenue','migratedAtRiskToActive','migratedActiveToAtRisk','reactivatedFromLapsed','avgOrdersPerActiveCustomer'], sample: [['2025-03','620','4100','1080','2500','8300','3.3','29800','131200','10800','2500','175','70','40','2.4']] },
  outreach: { headers: ['week','channel','sends','delivered','responses','responseRate','conversions','conversionRate','revenue','cost'], sample: [['2025-03-24','WhatsApp','4500','4365','1833','42.0','200','4.58','14800','450']] },
  beforeAfter: { headers: ['activity','launchDate','metric','beforeValue','afterValue','beforePeriod','afterPeriod','lift','unit'], sample: [['Loyalty Program Launch','2024-10-01','Monthly Retention Rate','78.0','92.4','Jul-Sep 2024','Mar 2025','18.5','percent']] },
  holdoutTests: { headers: ['testName','testPeriod','sampleSize','controlSize','exposedSize','controlConversionRate','exposedConversionRate','controlRevPerCustomer','exposedRevPerCustomer','incrementalConversionLift','incrementalRevLift','incrementalRevenue','confidence','status'], sample: [['Abandoned Cart Flow','Jan-Mar 2025','42000','4200','37800','8.2','16.0','18.50','38.20','95.1','106.5','48200','0.96','active']] },
  activityROI: { headers: ['activity','channel','totalCost','attributedRevenue','incrementalRevenue','incrementalROI','customersInfluenced','period'], sample: [['Welcome Series Flow','Email','480','66000','32400','66.5','12600','Q1 2025']] },
  revenue: { headers: ['week','totalRevenue','subscriptionRevenue','oneTimeRevenue','refunds','netRevenue','totalOrders','aov'], sample: [['2025-03-24','175200','128400','46800','2800','172400','2250','77.87']] },
  subscriptions: { headers: ['month','activeSubscribers','newSubscribers','churnedSubscribers','reactivated','mrr','churnRate','voluntaryChurn','involuntaryChurn','ltv','skipCount'], sample: [['2025-03','5520','620','420','40','125800','7.6','5.1','2.5','272','205']] },
  milestoneProducts: { headers: ['month','product','tier2nd','tier2ndAOV','tier2ndLTV','tier3rd','tier3rdAOV','tier3rdLTV','tier6th','tier6thAOV','tier6thLTV'], sample: [['2025-03','Stress & Anxiety','500','85.20','170.40','180','95.80','287.40','82','110.20','661.20']] },
  whatsappFlows: { headers: ['week','flowName','sends','delivered','responses','conversions','revenue','cost'], sample: [['2025-03-24','Welcome Message','1420','1377','551','70','4900','142']] },
  postcardFlows: { headers: ['week','flowName','sends','delivered','responses','conversions','revenue','cost'], sample: [['2025-03-24','Welcome Pack','740','703','49','24','2160','592']] },
  channelCosts: { headers: ['month','channel','cost','notes'], sample: [['2025-03','Emails','1300','Klaviyo subscription + campaigns']] },
  productChurn: { headers: ['month','product','activeSubscribers','churnedSubscribers','churnRate','voluntaryChurn','involuntaryChurn','newSubscribers','reactivated'], sample: [['2025-03','Adult Dry Food','2002','118','5.9','4.1','1.8','248','25']] },
};

// ─── ALERT GENERATION ───
function generateAlerts(state) {
  const alerts = [];
  const latestCampaign = [...state.emailFlows].reverse().find(r => r.type === 'Campaign');
  if (latestCampaign) {
    if (latestCampaign.openRate < 35) alerts.push({ severity: 'danger', metric: 'Email Open Rate', value: latestCampaign.openRate + '%', message: 'Below 35% — investigate deliverability' });
    else if (latestCampaign.openRate < 40) alerts.push({ severity: 'warning', metric: 'Email Open Rate', value: latestCampaign.openRate + '%', message: 'Approaching concern threshold' });
    if (latestCampaign.unsubRate > 0.5) alerts.push({ severity: 'danger', metric: 'Unsubscribe Rate', value: latestCampaign.unsubRate + '%', message: 'Above 0.5% — review frequency' });
  }
  const segs = state.segments;
  if (segs.length >= 2) {
    const curr = segs[segs.length - 1], prev = segs[segs.length - 2];
    const growth = ((curr.segAtRisk - prev.segAtRisk) / prev.segAtRisk) * 100;
    if (growth > 10) alerts.push({ severity: 'danger', metric: 'At-Risk Segment', value: formatNumber(curr.segAtRisk), message: `Grew ${growth.toFixed(1)}% MoM` });
    else if (curr.segAtRisk > 1200) alerts.push({ severity: 'warning', metric: 'At-Risk Segment', value: formatNumber(curr.segAtRisk), message: 'Above 1,200 — monitor closely' });
  }
  const loy = state.loyalty;
  if (loy.length > 0) {
    const latest = loy[loy.length - 1];
    if (latest.redemptionRate < 10) alerts.push({ severity: 'warning', metric: 'Milestone Redemption', value: latest.redemptionRate + '%', message: 'Below 10% — members not engaging' });
    const cumIssued = loy.reduce((s, m) => s + m.pointsIssued, 0);
    const cumRedeemed = loy.reduce((s, m) => s + m.pointsRedeemed, 0);
    const liability = (cumIssued - cumRedeemed) / 100;
    if (liability > 20000) alerts.push({ severity: 'warning', metric: 'Points Liability', value: formatCurrency(liability), message: 'Consider expiry policy' });
    if (latest.tier6thOrderLTV && latest.nonMemberLTV && latest.nonMemberLTV > 0) {
      const ltvLiftPct = ((latest.tier6thOrderLTV - latest.nonMemberLTV) / latest.nonMemberLTV) * 100;
      if (ltvLiftPct < 200) alerts.push({ severity: 'warning', metric: '6th Order LTV Lift', value: `${ltvLiftPct.toFixed(0)}%`, message: 'Below 200% vs non-members — review program incentives' });
    }
  }
  const lowConf = state.holdoutTests.filter(t => t.confidence < 0.80);
  if (lowConf.length > 0) alerts.push({ severity: 'warning', metric: 'Test Confidence', value: lowConf.length + ' test(s)', message: lowConf.map(t => t.testName).join(', ') + ' below 80%' });
  const weeks = [...new Set(state.emailFlows.filter(r => r.type === 'Flow').map(r => r.week))].sort();
  if (weeks.length >= 2) {
    const revByWeek = w => state.emailFlows.filter(r => r.week === w && r.type === 'Flow').reduce((s, r) => s + r.revenue, 0);
    const latest = revByWeek(weeks[weeks.length - 1]), prev = revByWeek(weeks[weeks.length - 2]);
    const drop = prev > 0 ? ((latest - prev) / prev) * 100 : 0;
    if (drop < -15) alerts.push({ severity: 'danger', metric: 'Flow Revenue', value: formatCurrency(latest), message: `Dropped ${Math.abs(drop).toFixed(1)}% WoW` });
  }
  if (alerts.length === 0) alerts.push({ severity: 'good', metric: 'All Clear', value: '', message: 'All CRM KPIs within acceptable ranges' });
  return alerts;
}

// ─── REUSABLE COMPONENTS ───
function ChartTooltip({ active, payload, label, formatter }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: C.cardBg, border: `1px solid ${C.cardBorder}`, borderRadius: 4, padding: '10px 14px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
      <p style={{ margin: 0, fontWeight: 600, color: C.textPrimary, fontSize: 12 }}>{label}</p>
      {payload.map((p, i) => (
        <p key={i} style={{ margin: '4px 0 0', color: p.color || C.textSecondary, fontSize: 12 }}>
          {p.name}: {formatter ? formatter(p.value, p.name) : (typeof p.value === 'number' ? p.value.toLocaleString('en-GB') : p.value)}
        </p>
      ))}
    </div>
  );
}

function ChartHeader({ title, tooltip }) {
  const [show, setShow] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!show) return;
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setShow(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [show]);
  return (
    <div ref={ref} style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '0 0 16px', position: 'relative' }}>
      <h3 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: C.textPrimary }}>{title}</h3>
      {tooltip && (
        <span
          onClick={(e) => { e.stopPropagation(); setShow(s => !s); }}
          style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 18, height: 18, borderRadius: '50%', background: show ? '#272D45' : '#E5E5EB', color: show ? '#fff' : '#676986', fontSize: 11, fontWeight: 700, cursor: 'pointer', flexShrink: 0, userSelect: 'none', transition: 'all 0.15s' }}
        >?</span>
      )}
      {show && tooltip && (
        <div style={{ position: 'absolute', left: 0, top: '100%', marginTop: 6, background: '#272D45', color: '#fff', padding: '10px 14px', borderRadius: 6, fontSize: 12, lineHeight: 1.5, maxWidth: 320, zIndex: 50, boxShadow: '0 4px 16px rgba(0,0,0,0.18)', whiteSpace: 'pre-line' }}>
          {tooltip}
        </div>
      )}
    </div>
  );
}

function KPICard({ label, value, format = 'number', delta, compDelta, compLabel, status, sparkData, sparkKey, presentationMode }) {
  const fmt = (v) => {
    if (format === 'currency') return formatCurrency(v);
    if (format === 'currencyDecimal') return formatCurrencyDecimal(v);
    if (format === 'percent') return formatPercent(v);
    if (format === 'multiplier') return formatMultiplier(v);
    if (format === 'text') return v;
    return formatNumber(v);
  };
  const statusColor = status === 'good' ? C.success : status === 'warning' ? C.warning : status === 'danger' ? C.danger : C.textTertiary;
  const showComp = compDelta != null && compLabel;
  return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: presentationMode ? '20px' : '16px', border: `1px solid ${C.cardBorder}`, display: 'flex', flexDirection: 'column', gap: 4 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span style={{ fontSize: presentationMode ? 13 : 11, color: C.textSecondary, fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.05em' }}>{label}</span>
        <span style={{ width: 8, height: 8, borderRadius: '50%', background: statusColor }} />
      </div>
      <span style={{ fontSize: presentationMode ? 28 : 22, fontWeight: 700, color: C.textPrimary, lineHeight: 1.1 }}>{fmt(value)}</span>
      {delta != null && !showComp && (
        <span style={{ fontSize: 12, color: delta >= 0 ? C.success : C.danger, fontWeight: 600 }}>
          {delta >= 0 ? '▲' : '▼'} {Math.abs(delta).toFixed(1)}%
        </span>
      )}
      {showComp && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          <span style={{ fontSize: 12, fontWeight: 600, color: compDelta >= 0 ? C.success : C.danger }}>
            {compDelta >= 0 ? '▲' : '▼'} {Math.abs(compDelta).toFixed(1)}%
          </span>
          <span style={{ fontSize: 10, color: C.textTertiary }}>{compLabel}</span>
        </div>
      )}
      {sparkData?.length > 1 && sparkKey && (
        <div style={{ height: 32, marginTop: 4 }}>
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={sparkData}>
              <defs><linearGradient id={`sp-${sparkKey}`} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.primary} stopOpacity={0.3}/><stop offset="100%" stopColor={C.primary} stopOpacity={0}/></linearGradient></defs>
              <Area type="monotone" dataKey={sparkKey} stroke={C.primary} fill={`url(#sp-${sparkKey})`} strokeWidth={1.5} dot={false} />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function TabNav({ tabs, active, onSelect }) {
  return (
    <div className="crm-tab-nav" style={{ display: 'flex', gap: 4, overflowX: 'auto', padding: '8px 0', position: 'sticky', top: 0, zIndex: 10, background: C.pageBg, borderBottom: `1px solid ${C.cardBorder}`, WebkitOverflowScrolling: 'touch', msOverflowStyle: 'none', scrollbarWidth: 'none' }}>
      {tabs.map(t => (
        <button key={t.key} onClick={() => onSelect(t.key)} style={{ padding: '8px 16px', borderRadius: 4, border: 'none', cursor: 'pointer', fontSize: 13, fontWeight: active === t.key ? 700 : 500, background: active === t.key ? C.primary : 'transparent', color: active === t.key ? '#fff' : C.textSecondary, whiteSpace: 'nowrap', transition: 'all 0.15s', flexShrink: 0 }}>
          {t.label}
        </button>
      ))}
    </div>
  );
}

function CSVUploader({ label, source, requiredHeaders, dispatch }) {
  const fileRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [status, setStatus] = useState(null);
  const handleFile = useCallback((file) => {
    if (!file) return;
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: (results) => {
        const headers = results.meta.fields || [];
        const missing = requiredHeaders.filter(h => !headers.includes(h));
        if (missing.length > 0) { setStatus({ ok: false, msg: `Missing columns: ${missing.join(', ')}` }); return; }
        dispatch({ type: 'LOAD_DATA', source, payload: results.data });
        setStatus({ ok: true, msg: `Loaded ${results.data.length} rows` });
      },
      error: (err) => setStatus({ ok: false, msg: err.message }),
    });
  }, [requiredHeaders, source, dispatch]);
  const downloadTemplate = useCallback(() => {
    const t = CSV_TEMPLATES[source];
    if (!t) return;
    const csv = [t.headers.join(','), ...t.sample.map(r => r.join(','))].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${source}_template.csv`; a.click();
  }, [source]);
  return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.cardBorder}` }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <span style={{ fontWeight: 600, color: C.textPrimary, fontSize: 14 }}>{label}</span>
        <button onClick={downloadTemplate} style={{ fontSize: 12, color: C.primary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline' }}>Download Template</button>
      </div>
      <div onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
        onClick={() => fileRef.current?.click()} style={{ border: `2px dashed ${dragOver ? C.primary : C.cardBorder}`, borderRadius: 4, padding: '20px', textAlign: 'center', cursor: 'pointer', background: dragOver ? '#E8F0EB' : C.divider, transition: 'all 0.15s' }}>
        <p style={{ margin: 0, fontSize: 13, color: C.textSecondary }}>Drop CSV here or click to upload</p>
        <input ref={fileRef} type="file" accept=".csv" style={{ display: 'none' }} onChange={e => handleFile(e.target.files?.[0])} />
      </div>
      {status && <p style={{ marginTop: 8, fontSize: 12, color: status.ok ? C.success : C.danger }}>{status.msg}</p>}
    </div>
  );
}

// ─── SETTINGS MODAL ───
// ─── AUTH HELPERS ───
async function hashPassword(pw) {
  const enc = new TextEncoder().encode(pw);
  const hash = await crypto.subtle.digest('SHA-256', enc);
  return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
}

function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState('login');
  const [username, setUsername] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPw, setConfirmPw] = useState('');
  const [connStr, setConnStr] = useState(() => getNeonConnection());
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [needsSetup, setNeedsSetup] = useState(!getNeonConnection());

  const handleConnect = async () => {
    if (!connStr.trim()) { setError('Please enter a Neon connection string.'); return; }
    setLoading(true); setError(null);
    try {
      const sql = neon(connStr);
      await sql`SELECT 1`;
      localStorage.setItem('crm_neon_connection', connStr);
      await initNeonTables(connStr);
      setNeedsSetup(false);
    } catch (e) { setError('Connection failed: ' + e.message); }
    finally { setLoading(false); }
  };

  const handleLogin = async () => {
    if (!username.trim() || !password) { setError('Please fill in all fields.'); return; }
    setLoading(true); setError(null);
    try {
      const conn = getNeonConnection();
      const hash = await hashPassword(password);
      const rows = await neonQuery(conn, 'SELECT * FROM users WHERE username = $1', [username.trim().toLowerCase()]);
      if (rows.length === 0) { setError('User not found. Create an account first.'); setLoading(false); return; }
      if (rows[0].password_hash !== hash) { setError('Incorrect password.'); setLoading(false); return; }
      const user = rows[0];
      localStorage.setItem('crm_user_id', String(user.id));
      localStorage.setItem('crm_username', user.username);
      localStorage.setItem('crm_display_name', user.display_name);
      localStorage.setItem('crm_user_role', user.role || 'user');
      onLogin({ id: user.id, username: user.username, displayName: user.display_name, role: user.role || 'user' });
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  };

  const handleRegister = async () => {
    if (!username.trim() || !displayName.trim() || !password || !confirmPw) { setError('Please fill in all fields.'); return; }
    if (password !== confirmPw) { setError('Passwords do not match.'); return; }
    if (password.length < 4) { setError('Password must be at least 4 characters.'); return; }
    setLoading(true); setError(null);
    try {
      const conn = getNeonConnection();
      const hash = await hashPassword(password);
      const uname = username.trim().toLowerCase();
      const existing = await neonQuery(conn, 'SELECT id FROM users WHERE username = $1', [uname]);
      if (existing.length > 0) { setError('Username already taken.'); setLoading(false); return; }
      const allUsers = await neonQuery(conn, 'SELECT id FROM users LIMIT 1', []);
      const role = allUsers.length === 0 ? 'admin' : 'user';
      const rows = await neonQuery(conn, 'INSERT INTO users (username, password_hash, display_name, role) VALUES ($1, $2, $3, $4) RETURNING *', [uname, hash, displayName.trim(), role]);
      const user = rows[0];
      localStorage.setItem('crm_user_id', String(user.id));
      localStorage.setItem('crm_username', user.username);
      localStorage.setItem('crm_display_name', user.display_name);
      localStorage.setItem('crm_user_role', user.role);
      onLogin({ id: user.id, username: user.username, displayName: user.display_name, role: user.role });
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  };

  const inputStyle = { width: '100%', padding: '10px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 14, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' };
  const btnStyle = { width: '100%', padding: '12px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 14, fontWeight: 700, cursor: loading ? 'wait' : 'pointer', opacity: loading ? 0.7 : 1 };

  return (
    <div style={{ minHeight: '100vh', background: C.pageBg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }}>
      <div style={{ background: C.cardBg, borderRadius: 10, padding: 36, width: 400, maxWidth: '90vw', boxShadow: '0 8px 40px rgba(0,0,0,0.12)', border: `1px solid ${C.cardBorder}` }}>
        <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 8 }}><OmniPetLogo height={36} /></div>
        <h1 style={{ margin: '0 0 4px', fontSize: 22, fontWeight: 700, color: C.textPrimary, textAlign: 'center' }}>CRM Dashboard</h1>
        <p style={{ margin: '0 0 24px', fontSize: 13, color: C.textTertiary, textAlign: 'center' }}>
          {needsSetup ? 'Connect to your database to get started' : mode === 'login' ? 'Sign in to your account' : 'Create a new account'}
        </p>

        {needsSetup ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Neon Connection String</label>
              <input type="password" value={connStr} onChange={e => setConnStr(e.target.value)} placeholder="postgresql://user:pass@ep-xxx.neon.tech/neondb" style={inputStyle} />
            </div>
            {error && <p style={{ margin: 0, fontSize: 12, color: C.danger }}>{error}</p>}
            <button onClick={handleConnect} disabled={loading} style={btnStyle}>{loading ? 'Connecting...' : 'Connect Database'}</button>
          </div>
        ) : mode === 'login' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Username</label>
              <input type="text" value={username} onChange={e => setUsername(e.target.value)} placeholder="your.username" style={inputStyle} onKeyDown={e => e.key === 'Enter' && handleLogin()} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Password</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Enter password" style={inputStyle} onKeyDown={e => e.key === 'Enter' && handleLogin()} />
            </div>
            {error && <p style={{ margin: 0, fontSize: 12, color: C.danger }}>{error}</p>}
            <button onClick={handleLogin} disabled={loading} style={btnStyle}>{loading ? 'Signing in...' : 'Sign In'}</button>
            <p style={{ margin: '4px 0 0', fontSize: 12, color: C.textTertiary, textAlign: 'center' }}>
              Don't have an account? <button onClick={() => { setMode('register'); setError(null); }} style={{ background: 'none', border: 'none', color: C.primary, fontWeight: 600, cursor: 'pointer', fontSize: 12 }}>Create one</button>
            </p>
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Username</label>
              <input type="text" value={username} onChange={e => setUsername(e.target.value)} placeholder="Choose a username" style={inputStyle} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Display Name</label>
              <input type="text" value={displayName} onChange={e => setDisplayName(e.target.value)} placeholder="Your full name" style={inputStyle} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Password</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Choose a password" style={inputStyle} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Confirm Password</label>
              <input type="password" value={confirmPw} onChange={e => setConfirmPw(e.target.value)} placeholder="Confirm password" style={inputStyle} onKeyDown={e => e.key === 'Enter' && handleRegister()} />
            </div>
            {error && <p style={{ margin: 0, fontSize: 12, color: C.danger }}>{error}</p>}
            <button onClick={handleRegister} disabled={loading} style={btnStyle}>{loading ? 'Creating account...' : 'Create Account'}</button>
            <p style={{ margin: '4px 0 0', fontSize: 12, color: C.textTertiary, textAlign: 'center' }}>
              Already have an account? <button onClick={() => { setMode('login'); setError(null); }} style={{ background: 'none', border: 'none', color: C.primary, fontWeight: 600, cursor: 'pointer', fontSize: 12 }}>Sign in</button>
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

function SettingsModal({ open, onClose, currentUser }) {
  const [apiKey, setApiKey] = useState(() => getAnthropicKey());
  const [neonConn, setNeonConn] = useState(() => getNeonConnection());
  const hasEnvNeon = !!import.meta.env.VITE_NEON_CONNECTION;
  const hasEnvAnthropicKey = !!import.meta.env.VITE_ANTHROPIC_KEY;
  const [currentPw, setCurrentPw] = useState('');
  const [newPw, setNewPw] = useState('');
  const [confirmPw, setConfirmPw] = useState('');
  const [pwMsg, setPwMsg] = useState(null);
  const [pwLoading, setPwLoading] = useState(false);

  if (!open) return null;

  const save = () => {
    localStorage.setItem('crm_anthropic_key', apiKey);
    localStorage.setItem('crm_neon_connection', neonConn);
    onClose();
  };

  const handleChangePassword = async () => {
    if (!currentPw || !newPw || !confirmPw) { setPwMsg({ type: 'error', text: 'Please fill in all password fields.' }); return; }
    if (newPw.length < 4) { setPwMsg({ type: 'error', text: 'New password must be at least 4 characters.' }); return; }
    if (newPw !== confirmPw) { setPwMsg({ type: 'error', text: 'New passwords do not match.' }); return; }
    setPwLoading(true); setPwMsg(null);
    try {
      const conn = getNeonConnection();
      const currentHash = await hashPassword(currentPw);
      const rows = await neonQuery(conn, 'SELECT password_hash FROM users WHERE id = $1', [currentUser.id]);
      if (rows.length === 0 || rows[0].password_hash !== currentHash) {
        setPwMsg({ type: 'error', text: 'Current password is incorrect.' }); setPwLoading(false); return;
      }
      const newHash = await hashPassword(newPw);
      await neonQuery(conn, 'UPDATE users SET password_hash = $1 WHERE id = $2', [newHash, currentUser.id]);
      setCurrentPw(''); setNewPw(''); setConfirmPw('');
      setPwMsg({ type: 'success', text: 'Password updated successfully.' });
    } catch (e) { setPwMsg({ type: 'error', text: 'Error: ' + e.message }); }
    finally { setPwLoading(false); }
  };

  const inputStyle = { width: '100%', padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' };
  const labelStyle = { display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={onClose}>
      <div style={{ background: C.cardBg, borderRadius: 10, padding: 28, width: 460, maxWidth: '90vw', maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 20px 60px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        <h2 style={{ margin: '0 0 20px', fontSize: 18, fontWeight: 700, color: C.textPrimary }}>Settings</h2>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div>
            <label style={labelStyle}>Anthropic API Key (for AI Import)</label>
            {hasEnvAnthropicKey ? (
              <p style={{ margin: 0, fontSize: 12, color: C.success, fontWeight: 600 }}>Configured via environment variable</p>
            ) : (
              <input type="password" value={apiKey} onChange={e => setApiKey(e.target.value)} placeholder="sk-ant-..." style={inputStyle} />
            )}
          </div>
          <div>
            <label style={labelStyle}>Neon Database Connection String</label>
            {hasEnvNeon ? (
              <p style={{ margin: 0, fontSize: 12, color: C.success, fontWeight: 600 }}>Configured via environment variable</p>
            ) : (
              <>
                <input type="password" value={neonConn} onChange={e => setNeonConn(e.target.value)} placeholder="postgresql://user:pass@ep-xxx.neon.tech/neondb" style={inputStyle} />
                <p style={{ margin: '4px 0 0', fontSize: 11, color: C.textTertiary }}>Enables persistent initiative tracking with collaboration. Works without it using local demo data.</p>
              </>
            )}
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 16 }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
          <button onClick={save} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Save</button>
        </div>

        <div style={{ borderTop: `1px solid ${C.divider}`, margin: '20px 0 0', paddingTop: 20 }}>
          <h3 style={{ margin: '0 0 14px', fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Change Password</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div>
              <label style={labelStyle}>Current Password</label>
              <input type="password" value={currentPw} onChange={e => setCurrentPw(e.target.value)} placeholder="Enter current password" style={inputStyle} />
            </div>
            <div>
              <label style={labelStyle}>New Password</label>
              <input type="password" value={newPw} onChange={e => setNewPw(e.target.value)} placeholder="Enter new password" style={inputStyle} />
            </div>
            <div>
              <label style={labelStyle}>Confirm New Password</label>
              <input type="password" value={confirmPw} onChange={e => setConfirmPw(e.target.value)} placeholder="Confirm new password" style={inputStyle} onKeyDown={e => e.key === 'Enter' && handleChangePassword()} />
            </div>
            {pwMsg && <p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: pwMsg.type === 'error' ? C.danger : '#18917B' }}>{pwMsg.text}</p>}
            <button onClick={handleChangePassword} disabled={pwLoading} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: '#124A2B', color: '#fff', fontSize: 13, fontWeight: 600, cursor: pwLoading ? 'wait' : 'pointer', opacity: pwLoading ? 0.7 : 1, alignSelf: 'flex-start' }}>
              {pwLoading ? 'Updating...' : 'Update Password'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── USER MANAGEMENT MODAL (Admin only) ───
function UserManagementModal({ open, onClose, currentUser }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(false);
  const [resetPw, setResetPw] = useState({});
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [msg, setMsg] = useState(null);

  const conn = getNeonConnection();

  const loadUsers = useCallback(async () => {
    if (!open || !conn) return;
    setLoading(true);
    try {
      const rows = await neonQuery(conn, 'SELECT id, username, display_name, role, created_at FROM users ORDER BY created_at ASC', []);
      setUsers(rows);
    } catch (e) { console.error(e); }
    finally { setLoading(false); }
  }, [open, conn]);

  useEffect(() => { loadUsers(); }, [loadUsers]);

  const showMsg = (text) => { setMsg(text); setTimeout(() => setMsg(null), 3000); };

  const handleResetPassword = async (userId) => {
    const newPw = resetPw[userId];
    if (!newPw || newPw.length < 4) { showMsg('Password must be at least 4 characters.'); return; }
    try {
      const hash = await hashPassword(newPw);
      await neonQuery(conn, 'UPDATE users SET password_hash = $1 WHERE id = $2', [hash, userId]);
      setResetPw(prev => ({ ...prev, [userId]: '' }));
      showMsg('Password reset successfully.');
    } catch (e) { showMsg('Error: ' + e.message); }
  };

  const handleToggleRole = async (userId, currentRole) => {
    if (userId === currentUser.id) { showMsg('You cannot change your own role.'); return; }
    const newRole = currentRole === 'admin' ? 'user' : 'admin';
    try {
      await neonQuery(conn, 'UPDATE users SET role = $1 WHERE id = $2', [newRole, userId]);
      setUsers(prev => prev.map(u => u.id === userId ? { ...u, role: newRole } : u));
      showMsg(`Role changed to ${newRole}.`);
    } catch (e) { showMsg('Error: ' + e.message); }
  };

  const handleDeleteUser = async (userId) => {
    if (userId === currentUser.id) { showMsg('You cannot delete yourself.'); return; }
    try {
      await neonQuery(conn, 'DELETE FROM users WHERE id = $1', [userId]);
      setUsers(prev => prev.filter(u => u.id !== userId));
      setDeleteConfirm(null);
      showMsg('User deleted.');
    } catch (e) { showMsg('Error: ' + e.message); }
  };

  if (!open) return null;

  const overlay = { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center' };
  const modal = { background: C.cardBg, borderRadius: 10, padding: 28, width: 680, maxWidth: '94vw', maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 20px 60px rgba(0,0,0,0.2)' };
  const badge = (role) => ({
    display: 'inline-block', padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 700,
    background: role === 'admin' ? '#E8F0EB' : '#E8F5F0', color: role === 'admin' ? '#124A2B' : '#18917B'
  });
  const btnSm = (bg, color) => ({ padding: '4px 10px', borderRadius: 2, border: 'none', background: bg, color, fontSize: 11, fontWeight: 600, cursor: 'pointer' });

  return (
    <div style={overlay} onClick={onClose}>
      <div style={modal} onClick={e => e.stopPropagation()}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textPrimary }}>User Management</h2>
          <button onClick={onClose} style={{ background: 'none', border: 'none', fontSize: 20, cursor: 'pointer', color: C.textTertiary }}>&times;</button>
        </div>

        {msg && <div style={{ padding: '8px 14px', borderRadius: 4, background: '#E8F0EB', color: '#124A2B', fontSize: 13, fontWeight: 600, marginBottom: 16 }}>{msg}</div>}

        {loading ? <p style={{ color: C.textTertiary, fontSize: 13 }}>Loading users...</p> : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${C.divider}` }}>
                {['Display Name', 'Username', 'Role', 'Created', 'Actions'].map(h => (
                  <th key={h} style={{ padding: '8px 10px', textAlign: 'left', color: C.textSecondary, fontWeight: 600, fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.05em' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {users.map(u => (
                <React.Fragment key={u.id}>
                  <tr style={{ borderBottom: `1px solid ${C.divider}` }}>
                    <td style={{ padding: '10px' }}>
                      <span style={{ fontWeight: 600, color: C.textPrimary }}>{u.display_name}</span>
                      {u.id === currentUser.id && <span style={{ marginLeft: 6, fontSize: 10, color: C.textTertiary }}>(you)</span>}
                    </td>
                    <td style={{ padding: '10px', color: C.textSecondary }}>{u.username}</td>
                    <td style={{ padding: '10px' }}><span style={badge(u.role)}>{u.role}</span></td>
                    <td style={{ padding: '10px', color: C.textTertiary, fontSize: 12 }}>{u.created_at ? new Date(u.created_at).toLocaleDateString() : '—'}</td>
                    <td style={{ padding: '10px' }}>
                      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                        <button onClick={() => setResetPw(prev => ({ ...prev, [u.id]: prev[u.id] !== undefined ? undefined : '' }))} style={btnSm('#E8F0EB', '#124A2B')}>Reset PW</button>
                        {u.id !== currentUser.id && (
                          <>
                            <button onClick={() => handleToggleRole(u.id, u.role)} style={btnSm('#E8F5F0', '#18917B')}>{u.role === 'admin' ? 'Demote' : 'Promote'}</button>
                            <button onClick={() => setDeleteConfirm(u.id)} style={btnSm('#FDE8E8', '#D81F26')}>Delete</button>
                          </>
                        )}
                      </div>
                    </td>
                  </tr>
                  {resetPw[u.id] !== undefined && (
                    <tr style={{ background: C.pageBg }}>
                      <td colSpan={5} style={{ padding: '8px 10px' }}>
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                          <input
                            type="password"
                            placeholder="New password (min 4 chars)"
                            value={resetPw[u.id] || ''}
                            onChange={e => setResetPw(prev => ({ ...prev, [u.id]: e.target.value }))}
                            style={{ flex: 1, padding: '6px 10px', borderRadius: 6, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.cardBg, color: C.textPrimary }}
                          />
                          <button onClick={() => handleResetPassword(u.id)} style={btnSm(C.primary, '#fff')}>Confirm</button>
                          <button onClick={() => setResetPw(prev => ({ ...prev, [u.id]: undefined }))} style={btnSm('transparent', C.textTertiary)}>Cancel</button>
                        </div>
                      </td>
                    </tr>
                  )}
                  {deleteConfirm === u.id && (
                    <tr style={{ background: '#FDE8E8' }}>
                      <td colSpan={5} style={{ padding: '8px 10px' }}>
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                          <span style={{ fontSize: 12, color: '#D81F26', fontWeight: 600 }}>Delete {u.display_name}? This cannot be undone.</span>
                          <button onClick={() => handleDeleteUser(u.id)} style={btnSm('#D81F26', '#fff')}>Yes, Delete</button>
                          <button onClick={() => setDeleteConfirm(null)} style={btnSm('transparent', C.textTertiary)}>Cancel</button>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ─── TIME PERIOD TOGGLE ───
function TimePeriodToggle({ tab, tabPeriods, dispatch, options }) {
  const opts = options || ['weekly', 'monthly'];
  const labels = { weekly: 'Weekly', monthly: 'Monthly', all: 'All Time' };
  const current = tabPeriods[tab] || opts[0];
  return (
    <div style={{ display: 'flex', gap: 2, background: C.divider, borderRadius: 4, padding: 2, width: 'fit-content' }}>
      {opts.map(o => (
        <button key={o} onClick={() => dispatch({ type: 'SET_TAB_PERIOD', tab, period: o })} style={{ padding: '5px 14px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: current === o ? 700 : 500, background: current === o ? C.cardBg : 'transparent', color: current === o ? C.textPrimary : C.textSecondary, cursor: 'pointer', boxShadow: current === o ? '0 1px 3px rgba(0,0,0,0.1)' : 'none', transition: 'all 0.15s' }}>
          {labels[o]}
        </button>
      ))}
    </div>
  );
}

// ─── DATE RANGE PICKER ───
function DateRangePicker({ state, dispatch }) {
  const { dateRange, pendingDateRange, pendingComparison, pendingDateMode, datePickerOpen } = state;
  const [selectingStart, setSelectingStart] = useState(true);
  const [hoverDate, setHoverDate] = useState(null);
  const [viewMonth, setViewMonth] = useState(() => {
    const d = new Date(pendingDateRange.end + 'T00:00:00');
    return { year: d.getFullYear(), month: d.getMonth() - 1 }; // show 2 months ending with end date's month
  });
  const panelRef = useRef(null);

  useEffect(() => {
    if (!datePickerOpen) return;
    const handler = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) {
        dispatch({ type: 'CANCEL_DATE_RANGE' });
      }
    };
    setTimeout(() => document.addEventListener('mousedown', handler), 0);
    return () => document.removeEventListener('mousedown', handler);
  }, [datePickerOpen, dispatch]);

  const presets = getDatePresets();
  const compRange = computeComparisonRange(dateRange.start, dateRange.end, state.comparison);

  const handlePreset = (p) => {
    dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { start: p.start, end: p.end } });
    setSelectingStart(true);
    setHoverDate(null);
  };

  const handleDayClick = (dateStr) => {
    if (selectingStart) {
      dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { start: dateStr, end: dateStr } });
      setSelectingStart(false);
    } else {
      const start = pendingDateRange.start;
      if (dateStr < start) {
        dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { start: dateStr, end: start } });
      } else {
        dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { start, end: dateStr } });
      }
      setSelectingStart(true);
      setHoverDate(null);
    }
  };

  const prevMonths = () => setViewMonth(v => { const m = v.month - 1; return m < 0 ? { year: v.year - 1, month: 11 } : { year: v.year, month: m }; });
  const nextMonths = () => setViewMonth(v => { const m = v.month + 1; return m > 11 ? { year: v.year + 1, month: 0 } : { year: v.year, month: m }; });

  const renderMonth = (year, month) => {
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const startOffset = (firstDay + 6) % 7; // Monday start
    const cells = [];
    for (let i = 0; i < startOffset; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);

    return (
      <div style={{ flex: 1 }}>
        <div style={{ textAlign: 'center', fontWeight: 600, fontSize: 13, color: C.textPrimary, marginBottom: 8 }}>
          {monthNames[month]} {year}
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 1 }}>
          {['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'].map(d => (
            <div key={d} style={{ textAlign: 'center', fontSize: 10, fontWeight: 600, color: C.textTertiary, padding: '2px 0' }}>{d}</div>
          ))}
          {cells.map((day, i) => {
            if (!day) return <div key={`e${i}`} />;
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            const isStart = dateStr === pendingDateRange.start;
            const isEnd = dateStr === pendingDateRange.end;
            const effectiveEnd = !selectingStart && hoverDate ? (hoverDate >= pendingDateRange.start ? hoverDate : pendingDateRange.start) : pendingDateRange.end;
            const effectiveStart = !selectingStart && hoverDate && hoverDate < pendingDateRange.start ? hoverDate : pendingDateRange.start;
            const inRange = dateStr >= effectiveStart && dateStr <= effectiveEnd;
            const isHover = dateStr === hoverDate && !selectingStart;

            return (
              <div
                key={dateStr}
                onClick={() => handleDayClick(dateStr)}
                onMouseEnter={() => !selectingStart && setHoverDate(dateStr)}
                onMouseLeave={() => setHoverDate(null)}
                style={{
                  textAlign: 'center', padding: '5px 2px', fontSize: 12, cursor: 'pointer',
                  borderRadius: isStart || isEnd ? 4 : 0,
                  background: isStart || isEnd ? C.primary : inRange ? '#124A2B22' : isHover ? '#124A2B11' : 'transparent',
                  color: isStart || isEnd ? '#fff' : inRange ? C.primary : C.textPrimary,
                  fontWeight: isStart || isEnd ? 700 : 400,
                  transition: 'background 0.1s',
                }}
              >
                {day}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const month2 = viewMonth.month + 1 > 11 ? { year: viewMonth.year + 1, month: 0 } : { year: viewMonth.year, month: viewMonth.month + 1 };

  const activePreset = presets.find(p => p.start === pendingDateRange.start && p.end === pendingDateRange.end);

  const comparisons = [
    { value: 'none', label: 'No comparison' },
    { value: 'previous_period', label: 'Previous period' },
    { value: 'previous_month', label: 'Previous month' },
    { value: 'previous_year', label: 'Previous year' },
  ];

  // Trigger bar (always visible)
  const triggerBar = (
    <div
      onClick={() => dispatch({ type: 'TOGGLE_DATE_PICKER' })}
      style={{
        display: 'flex', alignItems: 'center', gap: 10, padding: '8px 14px', background: C.cardBg,
        border: `1px solid ${datePickerOpen ? C.primary : C.cardBorder}`, borderRadius: 6, cursor: 'pointer',
        transition: 'border-color 0.15s', width: 'fit-content', userSelect: 'none',
      }}
    >
      <span style={{ fontSize: 14 }}>&#128197;</span>
      <span style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>
        {formatDateDisplay(dateRange.start)} — {formatDateDisplay(dateRange.end)}
      </span>
      {state.comparison !== 'none' && compRange && (
        <span style={{ fontSize: 11, color: C.textTertiary, marginLeft: 4 }}>
          vs {formatDateDisplay(compRange.start)} — {formatDateDisplay(compRange.end)}
        </span>
      )}
      <span style={{ fontSize: 10, color: C.textTertiary, marginLeft: 4, transform: datePickerOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.15s' }}>&#9660;</span>
    </div>
  );

  if (!datePickerOpen) return triggerBar;

  return (
    <div style={{ position: 'relative' }}>
      {triggerBar}
      <div ref={panelRef} className="crm-dp-panel" style={{
        position: 'absolute', top: '100%', left: 0, marginTop: 4, zIndex: 100,
        background: C.cardBg, borderRadius: 8, border: `1px solid ${C.cardBorder}`,
        boxShadow: '0 8px 32px rgba(0,0,0,0.12)', display: 'flex', minWidth: 680,
      }}>
        {/* Presets sidebar */}
        <div className="crm-dp-presets" style={{ width: 150, borderRight: `1px solid ${C.divider}`, padding: '12px 0', display: 'flex', flexDirection: 'column', gap: 1 }}>
          <div style={{ padding: '0 12px 8px', fontSize: 11, fontWeight: 600, color: C.textTertiary, textTransform: 'uppercase' }}>Presets</div>
          {presets.map(p => (
            <button key={p.label} onClick={() => handlePreset(p)} style={{
              padding: '6px 12px', border: 'none', background: activePreset?.label === p.label ? '#124A2B11' : 'transparent',
              color: activePreset?.label === p.label ? C.primary : C.textSecondary, fontSize: 12, fontWeight: activePreset?.label === p.label ? 600 : 400,
              cursor: 'pointer', textAlign: 'left', borderLeft: activePreset?.label === p.label ? `3px solid ${C.primary}` : '3px solid transparent',
            }}>
              {p.label}
            </button>
          ))}
        </div>

        {/* Calendar area */}
        <div className="crm-dp-calendar" style={{ flex: 1, padding: 16 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12 }}>
            <button onClick={prevMonths} style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: 16, color: C.textSecondary, padding: '2px 8px' }}>&lsaquo;</button>
            <div className="crm-dp-months" style={{ display: 'flex', gap: 24, flex: 1, justifyContent: 'center' }}>
              {renderMonth(viewMonth.year, viewMonth.month)}
              {renderMonth(month2.year, month2.month)}
            </div>
            <button onClick={nextMonths} style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: 16, color: C.textSecondary, padding: '2px 8px' }}>&rsaquo;</button>
          </div>

          {/* Date inputs */}
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12 }}>
            <input
              type="date" value={pendingDateRange.start}
              onChange={e => dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { ...pendingDateRange, start: e.target.value } })}
              style={{ flex: 1, padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, color: C.textPrimary }}
            />
            <span style={{ fontSize: 12, color: C.textTertiary }}>to</span>
            <input
              type="date" value={pendingDateRange.end}
              onChange={e => dispatch({ type: 'SET_PENDING_DATE_RANGE', payload: { ...pendingDateRange, end: e.target.value } })}
              style={{ flex: 1, padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, color: C.textPrimary }}
            />
          </div>

          {/* Fixed/Rolling toggle */}
          <div style={{ display: 'flex', gap: 2, background: C.divider, borderRadius: 4, padding: 2, width: 'fit-content', marginBottom: 12 }}>
            {['fixed', 'rolling'].map(m => (
              <button key={m} onClick={() => dispatch({ type: 'SET_PENDING_DATE_MODE', payload: m })} style={{
                padding: '4px 14px', borderRadius: 4, border: 'none', fontSize: 11, fontWeight: pendingDateMode === m ? 700 : 500,
                background: pendingDateMode === m ? C.cardBg : 'transparent', color: pendingDateMode === m ? C.textPrimary : C.textSecondary,
                cursor: 'pointer', boxShadow: pendingDateMode === m ? '0 1px 3px rgba(0,0,0,0.1)' : 'none', textTransform: 'capitalize',
              }}>
                {m}
              </button>
            ))}
          </div>

          {/* Comparison */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: C.textTertiary, textTransform: 'uppercase', marginBottom: 6 }}>Compare</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              {comparisons.map(c => (
                <label key={c.value} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: 12, color: C.textSecondary }}>
                  <input
                    type="radio" name="comparison" value={c.value} checked={pendingComparison === c.value}
                    onChange={() => dispatch({ type: 'SET_PENDING_COMPARISON', payload: c.value })}
                    style={{ accentColor: C.primary }}
                  />
                  {c.label}
                </label>
              ))}
            </div>
          </div>

          {/* Footer */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, borderTop: `1px solid ${C.divider}`, paddingTop: 12 }}>
            <button onClick={() => dispatch({ type: 'CANCEL_DATE_RANGE' })} style={{ padding: '6px 16px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
            <button onClick={() => dispatch({ type: 'APPLY_DATE_RANGE' })} style={{ padding: '6px 16px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Apply</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── AI DATA IMPORTER ───
const DATASET_SCHEMAS = {
  emailFlows: { label: 'Email & Flows', fields: ['week','type','flowName','sends','delivered','opens','openRate','clicks','ctr','unsubscribes','unsubRate','revenue','conversions','listSize'], example: DEMO_EMAIL_FLOWS.slice(0, 3) },
  loyalty: { label: 'Milestone Reward', fields: ['month','totalMembers','newEnrollments','pointsIssued','pointsRedeemed','redemptionRate','rewardsRedeemed','rewardCostGBP','revenueFromMembers','revenueFromNonMembers','memberAOV','nonMemberAOV','memberRetentionRate','nonMemberRetentionRate','tier2ndOrderMembers','tier2ndOrderAOV','tier2ndOrderLTV','tier3rdOrderMembers','tier3rdOrderAOV','tier3rdOrderLTV','tier6thOrderMembers','tier6thOrderAOV','tier6thOrderLTV','nonMemberLTV'], example: DEMO_LOYALTY.slice(0, 2) },
  segments: { label: 'Segments & Lifecycle', fields: ['month','segNew','segActive','segAtRisk','segLapsed','totalCustomers','avgRFMScore','segNewRevenue','segActiveRevenue','segAtRiskRevenue','segLapsedRevenue','migratedAtRiskToActive','migratedActiveToAtRisk','reactivatedFromLapsed','avgOrdersPerActiveCustomer'], example: DEMO_SEGMENTS.slice(0, 2) },
  outreach: { label: 'Direct Outreach', fields: ['week','channel','sends','delivered','responses','responseRate','conversions','conversionRate','revenue','cost'], example: DEMO_OUTREACH.slice(0, 3) },
  beforeAfter: { label: 'Before/After Analysis', fields: ['activity','launchDate','metric','beforeValue','afterValue','beforePeriod','afterPeriod','lift','unit'], example: DEMO_BEFORE_AFTER.slice(0, 2) },
  holdoutTests: { label: 'Holdout Tests', fields: ['testName','testPeriod','sampleSize','controlSize','exposedSize','controlConversionRate','exposedConversionRate','controlRevPerCustomer','exposedRevPerCustomer','incrementalConversionLift','incrementalRevLift','incrementalRevenue','confidence','status'], example: DEMO_HOLDOUT_TESTS.slice(0, 2) },
  activityROI: { label: 'Activity ROI', fields: ['activity','channel','totalCost','attributedRevenue','incrementalRevenue','incrementalROI','customersInfluenced','period'], example: DEMO_ACTIVITY_ROI.slice(0, 2) },
  revenue: { label: 'Revenue', fields: ['week','totalRevenue','subscriptionRevenue','oneTimeRevenue','refunds','netRevenue','totalOrders','aov'], example: DEMO_REVENUE.slice(0, 2) },
  subscriptions: { label: 'Subscriptions', fields: ['month','activeSubscribers','newSubscribers','churnedSubscribers','reactivated','mrr','churnRate','voluntaryChurn','involuntaryChurn','ltv','skipCount'], example: DEMO_SUBSCRIPTIONS.slice(0, 2) },
  milestoneProducts: { label: 'Milestone Products', fields: ['month','product','tier2nd','tier2ndAOV','tier2ndLTV','tier3rd','tier3rdAOV','tier3rdLTV','tier6th','tier6thAOV','tier6thLTV'], example: DEMO_MILESTONE_PRODUCTS.slice(0, 3) },
  whatsappFlows: { label: 'WhatsApp Flows', fields: ['week','flowName','sends','delivered','responses','conversions','revenue','cost'], example: DEMO_WHATSAPP_FLOWS.slice(0, 3) },
  postcardFlows: { label: 'Postcard Flows', fields: ['week','flowName','sends','delivered','responses','conversions','revenue','cost'], example: DEMO_POSTCARD_FLOWS.slice(0, 3) },
  channelCosts: { label: 'Channel Costs', fields: ['month','channel','cost','notes'], example: DEMO_CHANNEL_COSTS.slice(0, 4) },
  productChurn: { label: 'Product Churn', fields: ['month','product','activeSubscribers','churnedSubscribers','churnRate','voluntaryChurn','involuntaryChurn','newSubscribers','reactivated'], example: DEMO_PRODUCT_CHURN.slice(0, 3) },
};

function AIDataImporter({ dispatch, onOpenSettings, onLogImport, dashboardState }) {
  const [selectedDataset, setSelectedDataset] = useState('emailFlows');
  const [rawInput, setRawInput] = useState('');
  const [inputMode, setInputMode] = useState('paste');
  const [loading, setLoading] = useState(false);
  const [preview, setPreview] = useState(null);
  const [summary, setSummary] = useState(null);
  const [error, setError] = useState(null);
  const [imageData, setImageData] = useState(null);
  const [imagePreviewUrl, setImagePreviewUrl] = useState(null);
  const [importDateStart, setImportDateStart] = useState(dashboardState?.dateRange?.start || '');
  const [importDateEnd, setImportDateEnd] = useState(dashboardState?.dateRange?.end || '');
  const fileRef = useRef(null);
  const imageRef = useRef(null);

  const apiKey = getAnthropicKey();
  const schema = DATASET_SCHEMAS[selectedDataset];

  const handleFileUpload = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => setRawInput(e.target.result);
    reader.readAsText(file);
  };

  const handleImageUpload = (file) => {
    if (!file) return;
    const validTypes = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
    if (!validTypes.includes(file.type)) { setError('Please upload a PNG, JPEG, GIF, or WebP image.'); return; }
    if (file.size > 20 * 1024 * 1024) { setError('Image must be under 20MB.'); return; }
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target.result.split(',')[1];
      setImageData({ base64, mediaType: file.type, name: file.name });
      setImagePreviewUrl(e.target.result);
      setError(null);
    };
    reader.readAsDataURL(file);
  };

  const clearImage = () => { setImageData(null); setImagePreviewUrl(null); if (imageRef.current) imageRef.current.value = ''; };

  const dateContext = (importDateStart || importDateEnd) ? `\n\nDate context provided by user: ${importDateStart ? `Start: ${importDateStart}` : ''}${importDateStart && importDateEnd ? ', ' : ''}${importDateEnd ? `End: ${importDateEnd}` : ''}. If the data does not include explicit dates, use this date range to generate appropriate weekly (YYYY-MM-DD Mondays) or monthly (YYYY-MM) date values spanning this range.` : '';
  const systemPrompt = `You are a data formatting assistant. Convert the user's raw data into a JSON object with two keys: "summary" and "data".\n\nFields: ${schema.fields.join(', ')}\n\nExample rows:\n${JSON.stringify(schema.example, null, 2)}${dateContext}\n\nRules:\n- Return ONLY a valid JSON object with exactly two keys: "summary" (string) and "data" (array)\n- No markdown, no explanation, no code fences — just the raw JSON object\n- "data" must be an array of objects matching the schema above\n- "summary" must be a brief human-readable description (2-4 sentences) of what you found: number of rows, date range if applicable, key totals (e.g. total revenue, total sends), and any data quality notes (missing fields, assumptions made)\n- Use the exact field names shown\n- Convert dates to the format shown in examples\n- Numeric fields should be numbers, not strings\n- If data is missing a field, use null\n- Ensure all rows have all fields`;

  const organizeWithAI = async () => {
    if (!apiKey) { setError('No API key configured. Please open Settings.'); return; }
    const hasInput = inputMode === 'screenshot' ? !!imageData : !!rawInput.trim();
    if (!hasInput) { setError(inputMode === 'screenshot' ? 'Please upload a screenshot first.' : 'Please paste data or upload a file first.'); return; }
    setLoading(true); setError(null); setPreview(null);
    try {
      let messages;
      if (inputMode === 'screenshot') {
        messages = [{ role: 'user', content: [
          { type: 'image', source: { type: 'base64', media_type: imageData.mediaType, data: imageData.base64 } },
          { type: 'text', text: `Extract the data from this screenshot and convert it into the ${schema.label} format. Read every row visible in the image carefully.` }
        ]}];
      } else {
        messages = [{ role: 'user', content: `Convert this data into the ${schema.label} format:\n\n${rawInput}` }];
      }
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json', 'anthropic-dangerous-direct-browser-access': 'true' },
        body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 4096, system: systemPrompt, messages }),
      });
      if (!resp.ok) { const err = await resp.json().catch(() => ({})); throw new Error(err.error?.message || `API error ${resp.status}`); }
      const data = await resp.json();
      const text = data.content?.[0]?.text || '';
      // Try parsing as {summary, data} object first, fall back to raw array
      const objMatch = text.match(/\{[\s\S]*\}/);
      if (objMatch) {
        try {
          const obj = JSON.parse(objMatch[0]);
          if (obj.data && Array.isArray(obj.data)) {
            setSummary(obj.summary || null);
            setPreview(obj.data);
            return;
          }
        } catch (_) { /* fall through to array parsing */ }
      }
      const jsonMatch = text.match(/\[[\s\S]*\]/);
      if (!jsonMatch) throw new Error('AI did not return valid JSON');
      const parsed = JSON.parse(jsonMatch[0]);
      setSummary(null);
      setPreview(parsed);
    } catch (err) {
      setError(err.message);
    } finally { setLoading(false); }
  };

  const importData = () => {
    if (!preview) return;
    dispatch({ type: 'LOAD_DATA', source: selectedDataset, payload: preview });
    if (onLogImport) onLogImport({ dataset: schema.label, datasetKey: selectedDataset, inputMode, rowCount: preview.length, summary: summary || `${preview.length} rows imported` });
    setPreview(null); setSummary(null); setRawInput(''); setImageData(null); setImagePreviewUrl(null); setError(null);
  };

  const hasInput = inputMode === 'screenshot' ? !!imageData : !!rawInput.trim();

  return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
      <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>AI-Powered Data Import</h3>
      {!apiKey && (
        <div style={{ background: '#FEF3C7', borderRadius: 4, padding: 12, marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 13, color: '#92400E' }}>No API key configured.</span>
          <button onClick={onOpenSettings} style={{ fontSize: 12, color: C.primary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline', fontWeight: 600 }}>Open Settings</button>
        </div>
      )}
      <div style={{ display: 'flex', gap: 12, marginBottom: 16, alignItems: 'center', flexWrap: 'wrap' }}>
        <select value={selectedDataset} onChange={e => { setSelectedDataset(e.target.value); setPreview(null); setSummary(null); setError(null); }} style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }}>
          {Object.entries(DATASET_SCHEMAS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <div style={{ display: 'flex', gap: 2, background: C.divider, borderRadius: 6, padding: 2 }}>
          {[['paste', 'Paste Text'], ['file', 'Upload File'], ['screenshot', 'Screenshot']].map(([m, label]) => (
            <button key={m} onClick={() => setInputMode(m)} style={{ padding: '5px 12px', borderRadius: 4, border: 'none', fontSize: 12, fontWeight: inputMode === m ? 700 : 500, background: inputMode === m ? C.cardBg : 'transparent', color: inputMode === m ? C.textPrimary : C.textSecondary, cursor: 'pointer' }}>
              {label}
            </button>
          ))}
        </div>
      </div>
      {/* Date range context for AI */}
      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12, padding: '10px 14px', background: '#F9F7F2', borderRadius: 6, border: `1px solid ${C.divider}` }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: C.textSecondary }}>📅 Date Range</span>
        <input type="date" value={importDateStart} onChange={e => setImportDateStart(e.target.value)} style={{ padding: '5px 8px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', color: C.textPrimary, fontFamily: 'inherit' }} />
        <span style={{ fontSize: 11, color: C.textTertiary }}>to</span>
        <input type="date" value={importDateEnd} onChange={e => setImportDateEnd(e.target.value)} style={{ padding: '5px 8px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', color: C.textPrimary, fontFamily: 'inherit' }} />
        <span style={{ fontSize: 10, color: C.textTertiary, fontStyle: 'italic' }}>Used by AI when data doesn't include dates</span>
      </div>
      {inputMode === 'paste' ? (
        <textarea value={rawInput} onChange={e => setRawInput(e.target.value)} placeholder="Paste your raw data here — CSV, tab-separated, JSON, or any format. The AI will organize it..." rows={6} style={{ width: '100%', padding: '10px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, fontFamily: 'monospace', resize: 'vertical', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' }} />
      ) : inputMode === 'file' ? (
        <div onClick={() => fileRef.current?.click()} style={{ border: `2px dashed ${C.cardBorder}`, borderRadius: 4, padding: 20, textAlign: 'center', cursor: 'pointer', background: C.divider }}>
          <p style={{ margin: 0, fontSize: 13, color: C.textSecondary }}>{rawInput ? `File loaded (${rawInput.length} chars)` : 'Click to upload CSV, TSV, or text file'}</p>
          <input ref={fileRef} type="file" accept=".csv,.tsv,.txt,.json" style={{ display: 'none' }} onChange={e => handleFileUpload(e.target.files?.[0])} />
        </div>
      ) : (
        <div style={{ border: `2px dashed ${imageData ? C.primary : C.cardBorder}`, borderRadius: 4, padding: imageData ? 12 : 20, textAlign: 'center', background: C.divider, transition: 'border-color 0.2s' }}>
          {imageData ? (
            <div style={{ position: 'relative' }}>
              <img src={imagePreviewUrl} alt="Screenshot preview" style={{ maxWidth: '100%', maxHeight: 250, borderRadius: 6, objectFit: 'contain' }} />
              <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8 }}>
                <span style={{ fontSize: 12, color: C.textSecondary }}>{imageData.name}</span>
                <button onClick={clearImage} style={{ fontSize: 11, color: C.danger, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline' }}>Remove</button>
              </div>
            </div>
          ) : (
            <div onClick={() => imageRef.current?.click()} style={{ cursor: 'pointer', padding: 10 }}>
              <div style={{ fontSize: 28, marginBottom: 6 }}>&#128247;</div>
              <p style={{ margin: '0 0 4px', fontSize: 13, fontWeight: 600, color: C.textSecondary }}>Upload a screenshot</p>
              <p style={{ margin: 0, fontSize: 11, color: C.textTertiary }}>PNG, JPEG, GIF, or WebP — tables, reports, dashboards, spreadsheets</p>
            </div>
          )}
          <input ref={imageRef} type="file" accept="image/png,image/jpeg,image/gif,image/webp" style={{ display: 'none' }} onChange={e => handleImageUpload(e.target.files?.[0])} />
        </div>
      )}
      <div style={{ display: 'flex', gap: 10, marginTop: 12, alignItems: 'center' }}>
        <button onClick={organizeWithAI} disabled={loading || !hasInput} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: loading ? C.textTertiary : C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: loading ? 'wait' : 'pointer', opacity: !hasInput ? 0.5 : 1 }}>
          {loading ? 'Organizing...' : 'Organize with AI'}
        </button>
        {error && <span style={{ fontSize: 12, color: C.danger }}>{error}</span>}
      </div>
      {summary && preview && (
        <div style={{ marginTop: 16, background: '#E8F0EB', borderRadius: 6, padding: 14, border: '1px solid #A8D5BA' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
            <span style={{ fontSize: 16 }}>&#9993;</span>
            <span style={{ fontSize: 13, fontWeight: 700, color: '#4338CA' }}>AI Summary</span>
          </div>
          <p style={{ margin: 0, fontSize: 13, color: '#3730A3', lineHeight: 1.5 }}>{summary}</p>
        </div>
      )}
      {preview && (
        <div style={{ marginTop: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: C.success }}>{preview.length} rows parsed successfully</span>
            <button onClick={importData} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.success, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Confirm & Import</button>
          </div>
          <div style={{ maxHeight: 200, overflow: 'auto', borderRadius: 4, border: `1px solid ${C.cardBorder}` }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
              <thead>
                <tr>{schema.fields.slice(0, 8).map(f => <th key={f} style={{ padding: '6px 8px', textAlign: 'left', borderBottom: `1px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 10, position: 'sticky', top: 0, background: C.cardBg }}>{f}</th>)}</tr>
              </thead>
              <tbody>
                {preview.slice(0, 10).map((row, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                    {schema.fields.slice(0, 8).map(f => <td key={f} style={{ padding: '4px 8px', color: C.textPrimary, fontSize: 11 }}>{row[f] != null ? String(row[f]).slice(0, 20) : '—'}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
            {preview.length > 10 && <p style={{ textAlign: 'center', fontSize: 11, color: C.textTertiary, padding: 4 }}>...and {preview.length - 10} more rows</p>}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── IMPORT ACTIVITY LOG ───
const MODE_LABELS = { paste: 'Text', file: 'File', screenshot: 'Screenshot' };
const DATASET_COLORS = { 'Email & Flows': '#124A2B', 'Milestone Reward': '#F59E0B', 'Segments & Lifecycle': '#18917B', 'Direct Outreach': '#2D8B6E', 'Before/After Analysis': '#676986', 'Holdout Tests': '#C0392B', 'Activity ROI': '#18917B', 'Revenue': '#D81F26', 'Subscriptions': '#E67E22' };

function timeAgo(dateStr) {
  const d = new Date(dateStr);
  const now = new Date();
  const diffMs = now - d;
  const mins = Math.floor(diffMs / 60000);
  if (mins < 1) return 'Just now';
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  const days = Math.floor(hrs / 24);
  if (days < 7) return `${days}d ago`;
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
}

function ImportActivityLog({ logs, onClear, currentUser }) {
  const [expandedId, setExpandedId] = useState(null);
  const [comments, setComments] = useState({});
  const [commentText, setCommentText] = useState('');
  const [commentCounts, setCommentCounts] = useState({});
  const connStr = getNeonConnection();

  useEffect(() => {
    if (!connStr || !logs || logs.length === 0) return;
    const ids = logs.filter(l => l.id).map(l => l.id);
    if (ids.length === 0) return;
    (async () => {
      try {
        const rows = await neonQuery(connStr, 'SELECT import_log_id, COUNT(*)::int as cnt FROM import_log_comments WHERE import_log_id = ANY($1) GROUP BY import_log_id', [ids]);
        const counts = {};
        rows.forEach(r => { counts[r.import_log_id] = r.cnt; });
        setCommentCounts(counts);
      } catch (_) {}
    })();
  }, [connStr, logs]);

  const loadComments = async (logId) => {
    if (!connStr || !logId) return;
    try {
      const rows = await neonQuery(connStr, 'SELECT * FROM import_log_comments WHERE import_log_id = $1 ORDER BY created_at ASC', [logId]);
      setComments(prev => ({ ...prev, [logId]: rows }));
    } catch (_) {}
  };

  const addComment = async (logId) => {
    if (!connStr || !commentText.trim() || !logId) return;
    const author = currentUser?.displayName || 'Anonymous';
    try {
      await neonQuery(connStr, 'INSERT INTO import_log_comments (import_log_id, author, content) VALUES ($1, $2, $3)', [logId, author, commentText.trim()]);
      setCommentText('');
      await loadComments(logId);
      setCommentCounts(prev => ({ ...prev, [logId]: (prev[logId] || 0) + 1 }));
    } catch (_) {}
  };

  const handleExpand = (key, logId) => {
    if (expandedId === key) { setExpandedId(null); return; }
    setExpandedId(key);
    if (logId && connStr) loadComments(logId);
  };

  if (!logs || logs.length === 0) return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
      <h3 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Import Activity Log</h3>
      <p style={{ margin: '12px 0 0', fontSize: 13, color: C.textTertiary }}>No imports yet. Use the AI importer above to get started.</p>
    </div>
  );

  return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <h3 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Import Activity Log</h3>
        {onClear && <button onClick={() => { if (confirm('Clear all import logs?')) onClear(); }} style={{ fontSize: 11, color: C.textTertiary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline' }}>Clear Log</button>}
      </div>
      <div style={{ maxHeight: 400, overflow: 'auto', display: 'flex', flexDirection: 'column', gap: 6 }}>
        {logs.map((log, i) => {
          const key = log.id || i;
          const isExpanded = expandedId === key;
          const logComments = comments[log.id] || [];
          const ccnt = commentCounts[log.id] || 0;
          return (
            <div key={key} style={{ borderRadius: 4, background: C.pageBg, border: `1px solid ${isExpanded ? C.primary + '44' : C.divider}`, transition: 'border-color 0.15s' }}>
              <div onClick={() => handleExpand(key, log.id)} style={{ padding: '10px 12px', cursor: 'pointer' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <span style={{ fontSize: 11, color: C.textTertiary, minWidth: 65 }}>{timeAgo(log.created_at)}</span>
                  <span style={{ fontSize: 11, fontWeight: 600, color: '#fff', background: DATASET_COLORS[log.dataset] || C.primary, padding: '2px 8px', borderRadius: 4 }}>{log.dataset}</span>
                  <span style={{ fontSize: 11, color: C.textSecondary, background: C.divider, padding: '2px 6px', borderRadius: 4 }}>{MODE_LABELS[log.input_mode] || log.input_mode}</span>
                  <span style={{ fontSize: 11, color: C.textSecondary }}>{log.row_count} rows</span>
                  {ccnt > 0 && <span style={{ fontSize: 10, color: C.primary, fontWeight: 700 }}>{ccnt} comment{ccnt !== 1 ? 's' : ''}</span>}
                  <span style={{ fontSize: 11, color: C.textTertiary, marginLeft: 'auto' }}>by {log.imported_by}</span>
                </div>
              </div>
              {isExpanded && (
                <div style={{ padding: '0 12px 12px', borderTop: `1px solid ${C.divider}` }}>
                  {log.summary && <p style={{ margin: '8px 0', fontSize: 12, color: C.textSecondary, lineHeight: 1.4 }}>{log.summary}</p>}
                  {/* Comments thread */}
                  <div style={{ marginTop: 8 }}>
                    {logComments.length > 0 && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 8 }}>
                        {logComments.map(c => (
                          <div key={c.id} style={{ display: 'flex', gap: 8, alignItems: 'flex-start' }}>
                            <div style={{ width: 24, height: 24, borderRadius: '50%', background: C.primary, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700, flexShrink: 0 }}>
                              {(c.author || '?')[0].toUpperCase()}
                            </div>
                            <div style={{ flex: 1 }}>
                              <div style={{ display: 'flex', gap: 6, alignItems: 'baseline' }}>
                                <span style={{ fontSize: 11, fontWeight: 700, color: C.textPrimary }}>{c.author}</span>
                                <span style={{ fontSize: 10, color: C.textTertiary }}>{new Date(c.created_at).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</span>
                              </div>
                              <p style={{ margin: '2px 0 0', fontSize: 12, color: C.textSecondary, lineHeight: 1.4 }}>{c.content}</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                    {connStr && log.id && (
                      <div style={{ display: 'flex', gap: 6 }} onClick={e => e.stopPropagation()}>
                        <input
                          type="text" value={commentText} onChange={e => setCommentText(e.target.value)}
                          placeholder="Add a comment..."
                          onKeyDown={e => { if (e.key === 'Enter' && commentText.trim()) addComment(log.id); }}
                          style={{ flex: 1, padding: '6px 10px', borderRadius: 6, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.cardBg, color: C.textPrimary, fontFamily: 'inherit' }}
                        />
                        <button onClick={() => addComment(log.id)} disabled={!commentText.trim()} style={{ padding: '6px 12px', borderRadius: 6, border: 'none', background: commentText.trim() ? C.primary : C.divider, color: commentText.trim() ? '#fff' : C.textTertiary, fontSize: 11, fontWeight: 600, cursor: commentText.trim() ? 'pointer' : 'default' }}>Post</button>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── NEON HELPERS ───
// Returns the Neon connection string: env var (build-time) → localStorage fallback
function getNeonConnection() {
  return import.meta.env.VITE_NEON_CONNECTION || localStorage.getItem('crm_neon_connection') || '';
}

// Returns the Anthropic API key: env var (build-time) → localStorage fallback
function getAnthropicKey() {
  return import.meta.env.VITE_ANTHROPIC_KEY || localStorage.getItem('crm_anthropic_key') || '';
}

async function neonQuery(connectionString, sqlText, params = []) {
  const sql = neon(connectionString);
  // Build tagged template call: neon requires tagged template syntax
  const parts = sqlText.split(/\$\d+/);
  const strings = Object.assign([...parts], { raw: [...parts] });
  return await sql(strings, ...params);
}

async function initNeonTables(connectionString) {
  const sql = neon(connectionString);
  await sql`CREATE TABLE IF NOT EXISTS initiatives (
    id SERIAL PRIMARY KEY,
    title TEXT NOT NULL,
    description TEXT,
    status TEXT DEFAULT 'To Do',
    priority TEXT DEFAULT 'Medium',
    owner TEXT,
    due_date DATE,
    category TEXT,
    created_by TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS initiative_comments (
    id SERIAL PRIMARY KEY,
    initiative_id INT REFERENCES initiatives(id) ON DELETE CASCADE,
    author TEXT NOT NULL,
    content TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS import_log (
    id SERIAL PRIMARY KEY,
    dataset TEXT NOT NULL,
    input_mode TEXT NOT NULL,
    row_count INT NOT NULL,
    summary TEXT,
    imported_by TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL,
    role TEXT DEFAULT 'user',
    created_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`ALTER TABLE users ADD COLUMN IF NOT EXISTS role TEXT DEFAULT 'user'`;
  await sql`CREATE TABLE IF NOT EXISTS import_log_comments (
    id SERIAL PRIMARY KEY,
    import_log_id INT REFERENCES import_log(id) ON DELETE CASCADE,
    author TEXT NOT NULL,
    content TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS segment_links (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    type TEXT DEFAULT 'segment',
    klaviyo_url TEXT,
    description TEXT,
    member_count INT DEFAULT 0,
    created_by TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS activity_log (
    id SERIAL PRIMARY KEY,
    action TEXT NOT NULL,
    category TEXT NOT NULL,
    detail TEXT,
    username TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW()
  )`;
  await sql`CREATE TABLE IF NOT EXISTS dashboard_data (
    id SERIAL PRIMARY KEY,
    data_key TEXT UNIQUE NOT NULL,
    data_value JSONB NOT NULL DEFAULT '[]',
    updated_at TIMESTAMPTZ DEFAULT NOW()
  )`;
}

// ─── NEON DATA PERSISTENCE ───
async function saveDashboardData(connStr, dataKey, dataValue) {
  if (!connStr) return;
  try {
    await neonQuery(connStr, 'INSERT INTO dashboard_data (data_key, data_value, updated_at) VALUES ($1, $2, NOW()) ON CONFLICT (data_key) DO UPDATE SET data_value = $2, updated_at = NOW()', [dataKey, JSON.stringify(dataValue)]);
  } catch (e) { console.warn('Failed to save dashboard data to Neon:', e); }
}

async function loadAllDashboardData(connStr) {
  if (!connStr) return null;
  try {
    const rows = await neonQuery(connStr, 'SELECT data_key, data_value FROM dashboard_data', []);
    if (!rows || rows.length === 0) return null;
    const result = {};
    rows.forEach(r => { result[r.data_key] = r.data_value; });
    return result;
  } catch (e) { console.warn('Failed to load dashboard data from Neon:', e); return null; }
}

// ─── ACTIVITY LOG COLORS ───
const CATEGORY_LOG_COLORS = {
  emailFlows: '#124A2B', loyalty: '#F59E0B', segments: '#18917B',
  outreach: '#2D8B6E', initiatives: '#676986', system: '#94A3B8',
  beforeAfter: '#676986', holdoutTests: '#C0392B', activityROI: '#18917B',
  revenue: '#D81F26', subscriptions: '#E67E22'
};

const SEGMENT_TYPE_COLORS = { segment: '#7C3AED', list: '#2563EB' };

// ─── INITIATIVES SECTION ───
const STATUS_COLORS = { 'To Do': '#94A3B8', 'In Progress': '#18917B', 'Done': '#124A2B', 'Blocked': '#D81F26' };
const PRIORITY_COLORS = { 'Low': '#94A3B8', 'Medium': '#18917B', 'High': '#F59E0B', 'Urgent': '#D81F26' };
const CATEGORIES = ['Email', 'Milestone Reward', 'Outreach', 'Segments', 'General'];
const STATUSES = ['To Do', 'In Progress', 'Done', 'Blocked'];
const PRIORITIES = ['Low', 'Medium', 'High', 'Urgent'];

const METRIC_GROUPS = [
  { key: 'emailFlows', label: 'Email & Flows', icon: '\u2709', metrics: [
    { key: 'totalEmailRevenue', label: 'Total Email Revenue', format: 'currency' },
    { key: 'flowRevenue', label: 'Flow Revenue', format: 'currency' },
    { key: 'campaignRevenue', label: 'Campaign Revenue', format: 'currency' },
    { key: 'avgOpenRate', label: 'Avg Open Rate', format: 'percent' },
    { key: 'avgCTR', label: 'Avg Click-Through Rate', format: 'percent' },
    { key: 'listSize', label: 'List Size', format: 'number' },
  ]},
  { key: 'loyalty', label: 'Milestone Reward', icon: '\u2B50', metrics: [
    { key: 'totalMembers', label: 'Total Members', format: 'number' },
    { key: 'newEnrollments', label: 'New Enrollments', format: 'number' },
    { key: 'redemptionRate', label: 'Redemption Rate', format: 'percent' },
    { key: 'memberAOV', label: 'Member AOV', format: 'currencyDecimal' },
    { key: 'aovLift', label: 'AOV Lift (Member vs Non)', format: 'percent' },
    { key: 'memberRetentionRate', label: 'Member Retention', format: 'percent' },
    { key: 'tier6thOrderLTV', label: '6th Order LTV', format: 'currencyDecimal' },
    { key: 'ltvLift', label: 'LTV Lift (6th vs Non)', format: 'percent' },
  ]},
  { key: 'segments', label: 'Segments & Lifecycle', icon: '\uD83D\uDCC8', metrics: [
    { key: 'totalCustomers', label: 'Total Customers', format: 'number' },
    { key: 'segActive', label: 'Active Customers', format: 'number' },
    { key: 'segAtRisk', label: 'At-Risk Customers', format: 'number' },
    { key: 'segLapsed', label: 'Lapsed Customers', format: 'number' },
    { key: 'avgRFMScore', label: 'Avg RFM Score', format: 'decimal' },
    { key: 'migratedAtRiskToActive', label: 'Rescued from At-Risk', format: 'number' },
  ]},
  { key: 'revenue', label: 'Revenue & Orders', icon: '\uD83D\uDCB7', metrics: [
    { key: 'totalRevenue', label: 'Total Revenue', format: 'currency' },
    { key: 'netRevenue', label: 'Net Revenue', format: 'currency' },
    { key: 'subscriptionRevenue', label: 'Subscription Revenue', format: 'currency' },
    { key: 'totalOrders', label: 'Total Orders', format: 'number' },
    { key: 'aov', label: 'AOV', format: 'currencyDecimal' },
  ]},
  { key: 'subscriptions', label: 'Subscriptions', icon: '\uD83D\uDD04', metrics: [
    { key: 'activeSubscribers', label: 'Active Subscribers', format: 'number' },
    { key: 'mrr', label: 'MRR', format: 'currency' },
    { key: 'churnRate', label: 'Churn Rate', format: 'percent' },
    { key: 'newSubscribers', label: 'New Subscribers', format: 'number' },
    { key: 'ltv', label: 'LTV', format: 'currency' },
  ]},
  { key: 'outreach', label: 'Outreach', icon: '\uD83D\uDCF1', metrics: [
    { key: 'outreachRevenue', label: 'Total Outreach Revenue', format: 'currency' },
    { key: 'outreachCost', label: 'Total Cost', format: 'currency' },
    { key: 'outreachROAS', label: 'ROAS', format: 'multiplier' },
    { key: 'waResponseRate', label: 'WhatsApp Response Rate', format: 'percent' },
    { key: 'smsConvRate', label: 'SMS Conversion Rate', format: 'percent' },
  ]},
  { key: 'incrementality', label: 'Incrementality', icon: '\uD83D\uDD2C', metrics: [
    { key: 'totalIncrementalRevenue', label: 'Total Incremental Revenue', format: 'currency' },
    { key: 'avgIncrementalROI', label: 'Avg Incremental ROI', format: 'multiplier' },
    { key: 'activeHoldoutTests', label: 'Active Holdout Tests', format: 'number' },
    { key: 'highestLiftActivity', label: 'Highest Lift Activity', format: 'text' },
  ]},
];

function InitiativesSection({ onOpenSettings, currentUser, presentationMode, onLogActivity, activityLog, dashboardState, dispatch }) {
  const connStr = getNeonConnection();
  const isLocal = !connStr;
  const username = currentUser?.displayName || 'Anonymous';
  const DEMO_INITIATIVES = [];
  const [initiatives, setInitiatives] = useState(isLocal ? DEMO_INITIATIVES : []);
  const [comments, setComments] = useState({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [statusFilter, setStatusFilter] = useState('All');
  const [catFilter, setCatFilter] = useState('All');
  const [showForm, setShowForm] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [commentText, setCommentText] = useState('');
  const [formData, setFormData] = useState({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' });
  const [dbReady, setDbReady] = useState(isLocal);
  const [nextLocalId, setNextLocalId] = useState(100);
  const [pitstopView, setPitstopView] = useState(false);

  // Calendar state
  const [calendarView, setCalendarView] = useState(false);
  const [calendarMonth, setCalendarMonth] = useState(() => { const now = new Date(); return { year: now.getFullYear(), month: now.getMonth() }; });
  const [calendarSelectedDate, setCalendarSelectedDate] = useState(null);

  // Presentation editor state
  const [showPresentationEditor, setShowPresentationEditor] = useState(false);
  const [presentationSlides, setPresentationSlides] = useState([]);
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const slideImageRef = useRef(null);

  // Metric picker state
  const [showMetricPicker, setShowMetricPicker] = useState(false);
  const [selectedMetricGroups, setSelectedMetricGroups] = useState(() =>
    METRIC_GROUPS.reduce((acc, g) => ({ ...acc, [g.key]: false }), { initiatives: true })
  );

  useEffect(() => {
    if (!connStr) return;
    const init = async () => {
      try {
        setLoading(true);
        try { await initNeonTables(connStr); } catch (_) { /* tables may already exist */ }
        setDbReady(true);
        await loadInitiatives();
      } catch (e) { setError(e.message); } finally { setLoading(false); }
    };
    init();
  }, [connStr]);

  const loadInitiatives = async () => {
    if (isLocal) return;
    try {
      const rows = await neonQuery(connStr, 'SELECT * FROM initiatives ORDER BY CASE priority WHEN \'Urgent\' THEN 0 WHEN \'High\' THEN 1 WHEN \'Medium\' THEN 2 ELSE 3 END, CASE status WHEN \'In Progress\' THEN 0 WHEN \'To Do\' THEN 1 WHEN \'Blocked\' THEN 2 ELSE 3 END, created_at DESC');
      setInitiatives(rows);
    } catch (e) { setError(e.message); }
  };

  const loadComments = async (initiativeId) => {
    if (isLocal) return;
    try {
      const rows = await neonQuery(connStr, 'SELECT * FROM initiative_comments WHERE initiative_id = $1 ORDER BY created_at ASC', [initiativeId]);
      setComments(prev => ({ ...prev, [initiativeId]: rows }));
    } catch (e) { setError(e.message); }
  };

  const saveInitiative = async () => {
    if (!formData.title.trim()) return;
    if (isLocal) {
      if (editItem) {
        setInitiatives(prev => prev.map(i => i.id === editItem.id ? { ...i, ...formData, due_date: formData.due_date || null } : i));
      } else {
        const newItem = { id: nextLocalId, ...formData, due_date: formData.due_date || null, created_by: username, created_at: new Date().toISOString() };
        setNextLocalId(n => n + 1);
        setInitiatives(prev => [newItem, ...prev]);
      }
      setShowForm(false); setEditItem(null);
      if (onLogActivity) onLogActivity({ action: editItem ? 'Initiative Updated' : 'Initiative Created', category: 'initiatives', detail: `${editItem ? 'Updated' : 'Created'}: ${formData.title}`, user: username });
      setFormData({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' });
      return;
    }
    try {
      if (editItem) {
        await neonQuery(connStr, 'UPDATE initiatives SET title=$1, description=$2, status=$3, priority=$4, owner=$5, due_date=$6, category=$7, updated_at=NOW() WHERE id=$8', [formData.title, formData.description, formData.status, formData.priority, formData.owner, formData.due_date || null, formData.category, editItem.id]);
      } else {
        await neonQuery(connStr, 'INSERT INTO initiatives (title, description, status, priority, owner, due_date, category, created_by) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)', [formData.title, formData.description, formData.status, formData.priority, formData.owner, formData.due_date || null, formData.category, username]);
      }
      setShowForm(false); setEditItem(null);
      if (onLogActivity) onLogActivity({ action: editItem ? 'Initiative Updated' : 'Initiative Created', category: 'initiatives', detail: `${editItem ? 'Updated' : 'Created'}: ${formData.title}`, user: username });
      setFormData({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' });
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const updateStatus = async (id, newStatus) => {
    if (isLocal) {
      setInitiatives(prev => prev.map(i => i.id === id ? { ...i, status: newStatus } : i));
      const item = initiatives.find(i => i.id === id);
      if (onLogActivity && item) onLogActivity({ action: 'Status Changed', category: 'initiatives', detail: `"${item.title}" → ${newStatus}`, user: username });
      return;
    }
    try {
      await neonQuery(connStr, 'UPDATE initiatives SET status=$1, updated_at=NOW() WHERE id=$2', [newStatus, id]);
      const item = initiatives.find(i => i.id === id);
      if (onLogActivity && item) onLogActivity({ action: 'Status Changed', category: 'initiatives', detail: `"${item.title}" → ${newStatus}`, user: username });
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const deleteInitiative = async (id) => {
    if (isLocal) {
      const item = initiatives.find(i => i.id === id);
      setInitiatives(prev => prev.filter(i => i.id !== id));
      if (onLogActivity && item) onLogActivity({ action: 'Initiative Deleted', category: 'initiatives', detail: `Deleted: ${item.title}`, user: username });
      return;
    }
    try {
      const item = initiatives.find(i => i.id === id);
      await neonQuery(connStr, 'DELETE FROM initiatives WHERE id=$1', [id]);
      if (onLogActivity && item) onLogActivity({ action: 'Initiative Deleted', category: 'initiatives', detail: `Deleted: ${item.title}`, user: username });
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const addComment = async (initiativeId) => {
    if (!commentText.trim()) return;
    if (isLocal) {
      const newComment = { id: Date.now(), initiative_id: initiativeId, author: username, content: commentText, created_at: new Date().toISOString() };
      setComments(prev => ({ ...prev, [initiativeId]: [...(prev[initiativeId] || []), newComment] }));
      setCommentText('');
      return;
    }
    try {
      await neonQuery(connStr, 'INSERT INTO initiative_comments (initiative_id, author, content) VALUES ($1,$2,$3)', [initiativeId, username, commentText]);
      setCommentText('');
      await loadComments(initiativeId);
    } catch (e) { setError(e.message); }
  };

  const toggleExpand = async (id) => {
    if (expandedId === id) { setExpandedId(null); return; }
    setExpandedId(id);
    if (!isLocal && !comments[id]) await loadComments(id);
  };

  /* Local mode works without Neon — no guard needed */

  const filtered = initiatives.filter(i => (statusFilter === 'All' || i.status === statusFilter) && (catFilter === 'All' || i.category === catFilter));
  const inputStyle = { padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary };

  // ─── Calendar helpers ───
  const renderCalendar = () => {
    const { year, month } = calendarMonth;
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const startPad = firstDay.getDay();
    const daysInMonth = lastDay.getDate();
    const monthName = firstDay.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
    const cells = [];
    for (let i = 0; i < startPad; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);
    const initByDate = {};
    initiatives.forEach(item => { if (!item.due_date) return; const dStr = item.due_date.slice(0, 10); if (!initByDate[dStr]) initByDate[dStr] = []; initByDate[dStr].push(item); });
    const dateStr = (day) => `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    const today = new Date().toISOString().slice(0, 10);
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <button onClick={() => setCalendarMonth(p => ({ year: p.month === 0 ? p.year - 1 : p.year, month: p.month === 0 ? 11 : p.month - 1 }))} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, cursor: 'pointer' }}>{'\u2190'} Prev</button>
          <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textPrimary }}>{monthName}</h3>
          <button onClick={() => setCalendarMonth(p => ({ year: p.month === 11 ? p.year + 1 : p.year, month: p.month === 11 ? 0 : p.month + 1 }))} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, cursor: 'pointer' }}>Next {'\u2192'}</button>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 1 }}>
          {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(d => (
            <div key={d} style={{ padding: 8, textAlign: 'center', fontSize: 12, fontWeight: 600, color: C.textSecondary, background: C.divider, borderRadius: 2 }}>{d}</div>
          ))}
        </div>
        <div className="crm-calendar-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 1 }}>
          {cells.map((day, idx) => {
            if (day === null) return <div key={`pad-${idx}`} style={{ minHeight: 90, background: '#F9FAFB', borderRadius: 4 }} />;
            const ds = dateStr(day);
            const items = initByDate[ds] || [];
            const isToday = ds === today;
            const isSelected = ds === calendarSelectedDate;
            return (
              <div key={ds} onClick={() => setCalendarSelectedDate(isSelected ? null : ds)}
                style={{ minHeight: 90, padding: 6, background: isSelected ? '#E8F0EB' : isToday ? '#FFFDE7' : C.cardBg, border: `1px solid ${isSelected ? C.primary : C.cardBorder}`, borderRadius: 4, cursor: 'pointer', transition: 'all 0.15s', overflow: 'hidden' }}>
                <div style={{ fontSize: 12, fontWeight: isToday ? 700 : 400, color: isToday ? C.primary : C.textSecondary, marginBottom: 4 }}>{day}</div>
                {items.slice(0, 3).map(item => (
                  <div key={item.id} className="crm-cal-item" style={{ fontSize: 10, padding: '2px 4px', marginBottom: 2, borderRadius: 3, background: STATUS_COLORS[item.status] + '20', borderLeft: `3px solid ${STATUS_COLORS[item.status]}`, color: C.textPrimary, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.title}</div>
                ))}
                {items.length > 3 && <div style={{ fontSize: 10, color: C.textTertiary }}>+{items.length - 3} more</div>}
                {items.length > 0 && <div className="crm-cal-count" style={{ display: 'none', fontSize: 10, color: C.primary, fontWeight: 600 }}>{items.length} item{items.length > 1 ? 's' : ''}</div>}
              </div>
            );
          })}
        </div>
        {calendarSelectedDate && (
          <div style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.primary}` }}>
            <h4 style={{ margin: '0 0 12px', fontSize: 14, fontWeight: 600, color: C.textPrimary }}>
              Initiatives due {new Date(calendarSelectedDate + 'T00:00:00').toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
            </h4>
            {(initByDate[calendarSelectedDate] || []).length === 0 && <p style={{ fontSize: 13, color: C.textTertiary }}>No initiatives due on this date.</p>}
            {(initByDate[calendarSelectedDate] || []).map(item => (
              <div key={item.id} style={{ padding: 10, marginBottom: 8, borderRadius: 6, border: `1px solid ${C.cardBorder}`, borderLeft: `4px solid ${STATUS_COLORS[item.status]}`, background: C.pageBg }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 6 }}>
                  <span style={{ fontWeight: 600, fontSize: 13, color: C.textPrimary }}>{item.title}</span>
                  <div style={{ display: 'flex', gap: 6 }}>
                    <span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, background: STATUS_COLORS[item.status], color: '#fff', fontWeight: 600 }}>{item.status}</span>
                    <span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, border: `1px solid ${PRIORITY_COLORS[item.priority]}`, color: PRIORITY_COLORS[item.priority], fontWeight: 600 }}>{item.priority}</span>
                  </div>
                </div>
                {item.description && <p style={{ margin: '6px 0 0', fontSize: 12, color: C.textSecondary }}>{item.description}</p>}
                <div style={{ marginTop: 6, display: 'flex', gap: 12, fontSize: 11, color: C.textTertiary }}>
                  {item.owner && <span>Owner: {item.owner}</span>}
                  <span>Category: {item.category}</span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  // ─── Presentation helpers ───
  const uid = () => `slide-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const formatMetricValue = (value, format) => {
    switch (format) {
      case 'currency': return formatCurrency(value);
      case 'currencyDecimal': return formatCurrencyDecimal(value);
      case 'percent': return formatPercent(value);
      case 'number': return formatNumber(value);
      case 'multiplier': return value.toFixed(2) + 'x';
      case 'decimal': return value.toFixed(1);
      case 'text': return value || 'N/A';
      default: return String(value);
    }
  };

  const updateSlide = (index, changes) => {
    setPresentationSlides(prev => prev.map((s, i) => i === index ? { ...s, ...changes } : s));
  };

  const moveSlide = (fromIdx, toIdx) => {
    setPresentationSlides(prev => {
      const arr = [...prev]; const [moved] = arr.splice(fromIdx, 1); arr.splice(toIdx, 0, moved); return arr;
    });
    if (activeSlideIndex === fromIdx) setActiveSlideIndex(toIdx);
    else if (fromIdx < toIdx && activeSlideIndex > fromIdx && activeSlideIndex <= toIdx) setActiveSlideIndex(activeSlideIndex - 1);
    else if (fromIdx > toIdx && activeSlideIndex >= toIdx && activeSlideIndex < fromIdx) setActiveSlideIndex(activeSlideIndex + 1);
  };

  const deleteSlide = (index) => {
    setPresentationSlides(prev => prev.filter((_, i) => i !== index));
    setActiveSlideIndex(prev => prev >= presentationSlides.length - 1 ? Math.max(0, presentationSlides.length - 2) : prev > index ? prev - 1 : prev);
  };

  const addNewSlide = () => {
    const newSlide = { id: uid(), type: 'custom', title: 'New Slide', content: '', bullets: [], image: null };
    setPresentationSlides(prev => [...prev, newSlide]);
    setActiveSlideIndex(presentationSlides.length);
  };

  const handleSlideImageUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const validTypes = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
    if (!validTypes.includes(file.type)) { alert('Please upload a PNG, JPEG, GIF, or WebP image.'); return; }
    if (file.size > 20 * 1024 * 1024) { alert('Image must be under 20MB.'); return; }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const base64 = dataUrl.split(',')[1];
      updateSlide(activeSlideIndex, { image: { base64, dataUrl, name: file.name, mediaType: file.type } });
    };
    reader.readAsDataURL(file);
    if (slideImageRef.current) slideImageRef.current.value = '';
  };

  const generatePresentationSlides = () => {
    const ds = dashboardState || {};
    const start = ds.dateRange?.start || (() => { const d = new Date(); d.setDate(d.getDate() - d.getDay()); return d.toISOString().slice(0, 10); })();
    const end = ds.dateRange?.end || new Date().toISOString().slice(0, 10);
    const comparison = ds.comparison || 'none';
    const fmt = d => new Date(d).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    const compLabel = comparison === 'previous_period' ? 'vs Previous Period' : comparison === 'previous_month' ? 'vs Previous Month' : comparison === 'previous_year' ? 'vs Previous Year' : '';

    // Compute current metric values
    const currentVals = computeMetricGroupValues(ds, start, end);

    // Compute comparison values if comparison mode is active
    let compVals = null;
    if (comparison !== 'none') {
      const compRange = computeComparisonRange(start, end, comparison);
      compVals = computeMetricGroupValues(ds, compRange.start, compRange.end);
    }

    const slides = [];

    // 1. Title slide
    slides.push({ id: uid(), type: 'title', title: 'Weekly CRM Pitstop', content: `${fmt(start)} \u2013 ${fmt(end)}${compLabel ? '\n' + compLabel : ''}\nomni.pet`, bullets: [], image: null, titleData: { dateRange: `${fmt(start)} \u2013 ${fmt(end)}`, compLabel, brand: 'omni.pet' } });

    // 2. Cross-tab metric slides
    METRIC_GROUPS.forEach(group => {
      if (!selectedMetricGroups[group.key]) return;
      const curGroup = currentVals[group.key] || {};
      const compGroup = compVals ? (compVals[group.key] || {}) : null;
      const metricData = [];
      const bullets = group.metrics.map(m => {
        const cur = curGroup[m.key];
        const formatted = formatMetricValue(cur, m.format);
        const entry = { key: m.key, label: m.label, value: cur, formatted, format: m.format, prevValue: null, prevFormatted: null, change: null, arrow: '\u2192' };
        if (!compGroup || m.format === 'text') { metricData.push(entry); return `${m.label}: ${formatted}`; }
        const prev = compGroup[m.key];
        if (prev === undefined || prev === null || m.format === 'text') { metricData.push(entry); return `${m.label}: ${formatted}`; }
        const prevFormatted = formatMetricValue(prev, m.format);
        const numCur = typeof cur === 'number' ? cur : 0;
        const numPrev = typeof prev === 'number' ? prev : 0;
        let arrow = '\u2192';
        let pctChange = '';
        let changeNum = 0;
        if (numPrev !== 0) {
          changeNum = ((numCur - numPrev) / Math.abs(numPrev)) * 100;
          if (changeNum > 0.5) { arrow = '\u2191'; pctChange = `+${changeNum.toFixed(1)}%`; }
          else if (changeNum < -0.5) { arrow = '\u2193'; pctChange = `${changeNum.toFixed(1)}%`; }
          else { pctChange = '0.0%'; }
        } else if (numCur > 0) { arrow = '\u2191'; pctChange = '+100%'; changeNum = 100; }
        entry.prevValue = prev; entry.prevFormatted = prevFormatted; entry.change = changeNum; entry.arrow = arrow;
        metricData.push(entry);
        return `${m.label}: ${formatted} (${compLabel}: ${prevFormatted} ${arrow}${pctChange})`;
      });
      slides.push({ id: uid(), type: 'metrics', title: `${group.icon} ${group.label}`, content: '', bullets, metricData, image: null });
    });

    // 3. Initiatives Summary slide (always)
    const totalCount = initiatives.length;
    const inProgressCount = initiatives.filter(i => i.status === 'In Progress').length;
    const doneCount = initiatives.filter(i => i.status === 'Done').length;
    const blockedCount = initiatives.filter(i => i.status === 'Blocked').length;
    const todoCount = initiatives.filter(i => i.status === 'To Do').length;
    const completionRate = totalCount > 0 ? ((doneCount / totalCount) * 100) : 0;
    const initiativeData = { total: totalCount, done: doneCount, inProgress: inProgressCount, blocked: blockedCount, todo: todoCount, completionRate, statusDistribution: [ { name: 'Done', value: doneCount, color: '124A2B' }, { name: 'In Progress', value: inProgressCount, color: '18917B' }, { name: 'To Do', value: todoCount, color: '94A3B8' }, { name: 'Blocked', value: blockedCount, color: 'D81F26' } ].filter(s => s.value > 0) };
    slides.push({ id: uid(), type: 'summary', title: 'Initiatives Summary', content: '', bullets: [`Total Initiatives: ${totalCount}`, `Completed: ${doneCount}`, `In Progress: ${inProgressCount}`, `Blocked: ${blockedCount}`, `To Do: ${todoCount}`, `Completion Rate: ${completionRate.toFixed(0)}%`], initiativeData, image: null });

    // 4. Activity Highlights slide
    const recentLogs = (activityLog || []).filter(l => { const d = new Date(l.timestamp); return d >= new Date(start) && d <= new Date(end); }).slice(0, 10);
    if (recentLogs.length > 0) {
      slides.push({ id: uid(), type: 'activity', title: 'Activity Highlights', content: '', bullets: recentLogs.map(l => `${l.action}: ${l.detail} (${l.user})`), image: null });
    }

    // 5. Per-category initiative slides
    CATEGORIES.forEach(cat => {
      const catItems = initiatives.filter(i => i.category === cat);
      if (catItems.length === 0) return;
      slides.push({ id: uid(), type: 'category', title: `${cat} Initiatives`, content: '', bullets: catItems.map(i => `[${i.status}] ${i.title}${i.owner ? ` (@${i.owner})` : ''}`), image: null });
    });

    // 6. Screenshot slide (always last)
    slides.push({ id: uid(), type: 'screenshot', title: 'Screenshots', content: 'Upload screenshots to this slide using the image upload button below.', bullets: [], image: null });

    setPresentationSlides(slides);
    setActiveSlideIndex(0);
    setShowMetricPicker(false);
    setShowPresentationEditor(true);
  };

  const downloadPptx = async () => {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'omni.pet CRM Dashboard';
    const BG = '124A2B';

    const renderBulletSlide = (s, slide) => {
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: BG } });
      s.addText(slide.title || 'Untitled', { x: 0.5, y: 0.1, w: 9, h: 0.6, fontSize: 24, fontFace: 'Arial', color: 'FFFFFF', bold: true });
      let yOff = 1.2;
      if (slide.content) { s.addText(slide.content, { x: 0.5, y: yOff, w: 9, h: 1, fontSize: 14, fontFace: 'Arial', color: '272D45', valign: 'top', wrap: true }); yOff += 1.2; }
      if (slide.bullets && slide.bullets.length > 0) {
        const bt = slide.bullets.filter(b => b.trim()).map(b => ({ text: b, options: { bullet: true, fontSize: 13, color: '272D45' } }));
        if (bt.length > 0) { const h = Math.min(4, bt.length * 0.4); s.addText(bt, { x: 0.5, y: yOff, w: 9, h, fontFace: 'Arial', valign: 'top' }); yOff += h + 0.3; }
      }
      if (slide.image) { const imgH = Math.min(4, 7.5 - yOff - 0.5); if (imgH > 0.5) s.addImage({ data: slide.image.dataUrl, x: 0.5, y: yOff, w: 8, h: imgH, sizing: { type: 'contain', w: 8, h: imgH } }); }
    };

    const renderTitleSlide = (s, slide) => {
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 4.0, fill: { color: BG } });
      s.addText(slide.title || 'Weekly CRM Pitstop', { x: 0.8, y: 1.0, w: 11, h: 1.2, fontSize: 40, fontFace: 'Arial', bold: true, color: 'FFFFFF' });
      if (slide.titleData?.dateRange) s.addText(slide.titleData.dateRange, { x: 0.8, y: 2.3, w: 11, h: 0.6, fontSize: 20, fontFace: 'Arial', color: 'D4E8DB' });
      if (slide.titleData?.compLabel) s.addText(slide.titleData.compLabel, { x: 0.8, y: 2.95, w: 11, h: 0.4, fontSize: 14, fontFace: 'Arial', italic: true, color: 'A8D5B8' });
      s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 4.6, w: 2, h: 0.04, fill: { color: BG } });
      s.addText('omni.pet', { x: 0.8, y: 4.9, w: 4, h: 0.6, fontSize: 24, fontFace: 'Arial', bold: true, color: BG });
    };

    const renderMetricsSlide = (s, slide) => {
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: BG } });
      s.addText(slide.title || 'Metrics', { x: 0.5, y: 0.1, w: 9, h: 0.6, fontSize: 24, fontFace: 'Arial', color: 'FFFFFF', bold: true });
      const colW = 3.9, rowH = 2.5, gapX = 0.25, gapY = 0.2, startX = 0.5, startY = 1.1;
      (slide.metricData || []).forEach((m, idx) => {
        if (idx >= 6) return;
        const col = idx % 3, row = Math.floor(idx / 3);
        const x = startX + col * (colW + gapX), y = startY + row * (rowH + gapY);
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: rowH, fill: { color: 'F8F9FA' }, line: { color: 'E5E5EB', width: 0.75 }, rectRadius: 0.1 });
        s.addText(m.formatted || '', { x: x + 0.3, y: y + 0.3, w: colW - 0.6, h: 0.7, fontSize: 28, fontFace: 'Arial', bold: true, color: '272D45', valign: 'top' });
        s.addText(m.label || '', { x: x + 0.3, y: y + 1.15, w: colW - 0.6, h: 0.35, fontSize: 11, fontFace: 'Arial', color: '676986', valign: 'top' });
        if (m.change !== null && m.change !== undefined) {
          const bc = m.arrow === '\u2191' ? '18917B' : m.arrow === '\u2193' ? 'D81F26' : '94A3B8';
          const bt = `${m.arrow} ${Math.abs(m.change).toFixed(1)}%`;
          s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 1.7, w: 1.6, h: 0.35, fill: { color: bc }, rectRadius: 0.15 });
          s.addText(bt, { x: x + 0.3, y: y + 1.7, w: 1.6, h: 0.35, fontSize: 10, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
          if (m.prevFormatted) s.addText(`vs ${m.prevFormatted}`, { x: x + 0.3, y: y + 2.1, w: colW - 0.6, h: 0.25, fontSize: 9, fontFace: 'Arial', color: '9CA3AF', valign: 'top' });
        }
      });
    };

    const renderSummarySlide = (s, slide) => {
      const d = slide.initiativeData;
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: BG } });
      s.addText(slide.title || 'Initiatives Summary', { x: 0.5, y: 0.1, w: 9, h: 0.6, fontSize: 24, fontFace: 'Arial', color: 'FFFFFF', bold: true });
      // Left: Donut chart
      if (d.statusDistribution && d.statusDistribution.length > 0) {
        try {
          s.addChart(pptx.charts.DOUGHNUT, [{ name: 'Status', labels: d.statusDistribution.map(x => x.name), values: d.statusDistribution.map(x => x.value) }], { x: 0.3, y: 1.1, w: 5.5, h: 4.5, holeSize: 60, showTitle: false, showValue: true, showPercent: false, showLegend: true, legendPos: 'b', legendFontSize: 10, chartColors: d.statusDistribution.map(x => x.color), dataLabelFontSize: 11, dataLabelColor: '272D45' });
        } catch(e) { /* fallback if chart fails */ }
      }
      // Right: 2x2 mini KPI cards
      const cards = [{ label: 'Total', value: String(d.total), color: '272D45' }, { label: 'Completed', value: String(d.done), color: '124A2B' }, { label: 'In Progress', value: String(d.inProgress), color: '18917B' }, { label: 'Blocked', value: String(d.blocked), color: 'D81F26' }];
      const mW = 3.0, mH = 1.8, mGX = 0.2, mGY = 0.2, mSX = 6.8, mSY = 1.1;
      cards.forEach((c, idx) => {
        const col = idx % 2, row = Math.floor(idx / 2);
        const mx = mSX + col * (mW + mGX), my = mSY + row * (mH + mGY);
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: mx, y: my, w: mW, h: mH, fill: { color: 'F8F9FA' }, line: { color: 'E5E5EB', width: 0.75 }, rectRadius: 0.08 });
        s.addText(c.value, { x: mx + 0.2, y: my + 0.2, w: mW - 0.4, h: 0.8, fontSize: 32, fontFace: 'Arial', bold: true, color: c.color, align: 'center' });
        s.addText(c.label, { x: mx + 0.2, y: my + 1.1, w: mW - 0.4, h: 0.4, fontSize: 11, fontFace: 'Arial', color: '676986', align: 'center' });
      });
      // Bottom: Completion progress bar
      const barY = 5.9, barX = 0.5, barW = 12.3, barH = 0.35, pct = d.completionRate / 100;
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: barX, y: barY, w: barW, h: barH, fill: { color: 'E5E7EB' }, rectRadius: 0.1 });
      if (pct > 0) s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: barX, y: barY, w: Math.max(0.3, barW * pct), h: barH, fill: { color: BG }, rectRadius: 0.1 });
      s.addText(`Completion: ${d.completionRate.toFixed(0)}%`, { x: barX, y: barY - 0.35, w: barW, h: 0.3, fontSize: 11, fontFace: 'Arial', bold: true, color: '272D45' });
    };

    presentationSlides.forEach(slide => {
      const s = pptx.addSlide();
      s.background = { color: 'FFFFFF' };
      switch (slide.type) {
        case 'title': renderTitleSlide(s, slide); break;
        case 'metrics': (slide.metricData && slide.metricData.length > 0) ? renderMetricsSlide(s, slide) : renderBulletSlide(s, slide); break;
        case 'summary': slide.initiativeData ? renderSummarySlide(s, slide) : renderBulletSlide(s, slide); break;
        default: renderBulletSlide(s, slide); break;
      }
      s.addText('omni.pet CRM Dashboard', { x: 0.5, y: 7.0, w: 4, h: 0.3, fontSize: 8, fontFace: 'Arial', color: '9CA3AF' });
    });
    const blob = await pptx.write({ outputType: 'blob' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `CRM_Pitstop_${new Date().toISOString().slice(0, 10)}.pptx`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {error && <div style={{ background: '#FDE8E8', borderRadius: 4, padding: 10, fontSize: 12, color: C.danger }}>{error} <button onClick={() => setError(null)} style={{ background: 'none', border: 'none', color: C.danger, cursor: 'pointer', fontWeight: 600, marginLeft: 8 }}>Dismiss</button></div>}

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          {['All', ...STATUSES].map(s => (
            <button key={s} onClick={() => setStatusFilter(s)} style={{ padding: '5px 14px', borderRadius: 4, border: `1px solid ${s === 'All' ? C.cardBorder : STATUS_COLORS[s] || C.cardBorder}`, background: statusFilter === s ? (s === 'All' ? C.textPrimary : STATUS_COLORS[s]) : 'transparent', color: statusFilter === s ? '#fff' : C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
              {s} {s !== 'All' && <span style={{ opacity: 0.7 }}>({initiatives.filter(i => i.status === s).length})</span>}
            </button>
          ))}
          <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ padding: '5px 10px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.pageBg, color: C.textSecondary }}>
            <option value="All">All Categories</option>
            {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button onClick={() => { setPitstopView(!pitstopView); if (!pitstopView) setCalendarView(false); }} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: pitstopView ? C.primary : 'transparent', color: pitstopView ? '#fff' : C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>
            {pitstopView ? 'Exit Pitstop' : 'Pitstop View'}
          </button>
          <button onClick={() => { setCalendarView(!calendarView); if (!calendarView) setPitstopView(false); }} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: calendarView ? C.primary : 'transparent', color: calendarView ? '#fff' : C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>
            {calendarView ? 'Exit Calendar' : 'Calendar View'}
          </button>
          <button onClick={() => { setShowForm(true); setEditItem(null); setFormData({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' }); }} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>+ Add Initiative</button>
        </div>
      </div>

      {showForm && (
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `2px solid ${C.primary}` }}>
          <h4 style={{ margin: '0 0 16px', fontSize: 14, fontWeight: 600, color: C.textPrimary }}>{editItem ? 'Edit Initiative' : 'New Initiative'}</h4>
          <div className="crm-2col-grid" style={{ display: 'grid', gap: 12 }}>
            <div style={{ gridColumn: '1/3' }}>
              <input value={formData.title} onChange={e => setFormData({ ...formData, title: e.target.value })} placeholder="Initiative title" style={{ ...inputStyle, width: '100%', boxSizing: 'border-box' }} />
            </div>
            <div style={{ gridColumn: '1/3' }}>
              <textarea value={formData.description} onChange={e => setFormData({ ...formData, description: e.target.value })} placeholder="Description (optional)" rows={2} style={{ ...inputStyle, width: '100%', resize: 'vertical', boxSizing: 'border-box' }} />
            </div>
            <select value={formData.status} onChange={e => setFormData({ ...formData, status: e.target.value })} style={inputStyle}>{STATUSES.map(s => <option key={s} value={s}>{s}</option>)}</select>
            <select value={formData.priority} onChange={e => setFormData({ ...formData, priority: e.target.value })} style={inputStyle}>{PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}</select>
            <input value={formData.owner} onChange={e => setFormData({ ...formData, owner: e.target.value })} placeholder="Owner" style={inputStyle} />
            <input type="date" value={formData.due_date} onChange={e => setFormData({ ...formData, due_date: e.target.value })} style={inputStyle} />
            <select value={formData.category} onChange={e => setFormData({ ...formData, category: e.target.value })} style={inputStyle}>{CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}</select>
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 16 }}>
            <button onClick={saveInitiative} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>{editItem ? 'Update' : 'Create'}</button>
            <button onClick={() => { setShowForm(false); setEditItem(null); }} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      )}

      {loading && <p style={{ textAlign: 'center', color: C.textSecondary, fontSize: 13 }}>Loading...</p>}

      {calendarView ? renderCalendar() : pitstopView ? (() => {
        const totalCount = initiatives.length;
        const inProgressCount = initiatives.filter(i => i.status === 'In Progress').length;
        const doneCount = initiatives.filter(i => i.status === 'Done').length;
        const blockedCount = initiatives.filter(i => i.status === 'Blocked').length;
        const todoCount = initiatives.filter(i => i.status === 'To Do').length;
        const completionPct = totalCount > 0 ? (doneCount / totalCount) * 100 : 0;
        const pFontSize = presentationMode ? 16 : 14;
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
            <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
              {[
                { label: 'Total Initiatives', value: totalCount, color: C.textPrimary },
                { label: 'In Progress', value: inProgressCount, color: STATUS_COLORS['In Progress'] },
                { label: 'Completed', value: doneCount, color: STATUS_COLORS['Done'] },
                { label: 'Blocked', value: blockedCount, color: STATUS_COLORS['Blocked'] },
                { label: 'Completion Rate', value: `${completionPct.toFixed(0)}%`, color: completionPct > 70 ? '#124A2B' : completionPct > 40 ? '#F59E0B' : '#D81F26' },
              ].map(card => (
                <div key={card.label} style={{ background: C.cardBg, borderRadius: 8, padding: presentationMode ? 28 : 20, border: `1px solid ${C.cardBorder}`, textAlign: 'center' }}>
                  <div style={{ fontSize: presentationMode ? 38 : 32, fontWeight: 700, color: card.color }}>{card.value}</div>
                  <div style={{ fontSize: pFontSize - 1, color: C.textSecondary, marginTop: 4 }}>{card.label}</div>
                </div>
              ))}
            </div>

            <div className="crm-pitstop-kanban" style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16 }}>
              {STATUSES.map(status => (
                <div key={status} style={{ background: C.pageBg, borderRadius: 8, padding: 12, border: `1px solid ${C.cardBorder}` }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
                    <div style={{ width: 10, height: 10, borderRadius: '50%', background: STATUS_COLORS[status] }} />
                    <span style={{ fontSize: pFontSize, fontWeight: 700, color: C.textPrimary }}>{status}</span>
                    <span style={{ fontSize: 12, color: C.textTertiary }}>({initiatives.filter(i => i.status === status).length})</span>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {initiatives.filter(i => i.status === status).map(item => (
                      <div key={item.id} style={{ background: C.cardBg, borderRadius: 6, padding: 14, border: `1px solid ${C.cardBorder}`, borderLeft: `4px solid ${PRIORITY_COLORS[item.priority]}` }}>
                        <div style={{ fontSize: pFontSize, fontWeight: 600, color: C.textPrimary, marginBottom: 6 }}>{item.title}</div>
                        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                          <span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, border: `1px solid ${PRIORITY_COLORS[item.priority]}`, color: PRIORITY_COLORS[item.priority], fontWeight: 600 }}>{item.priority}</span>
                          {item.owner && <span style={{ fontSize: 11, color: C.textSecondary }}>@{item.owner}</span>}
                          {item.due_date && <span style={{ fontSize: 11, color: new Date(item.due_date) < new Date() && item.status !== 'Done' ? '#D81F26' : C.textTertiary, fontWeight: new Date(item.due_date) < new Date() && item.status !== 'Done' ? 600 : 400 }}>{new Date(item.due_date).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}</span>}
                          <span style={{ fontSize: 10, padding: '2px 6px', borderRadius: 4, background: C.divider, color: C.textSecondary }}>{item.category}</span>
                        </div>
                      </div>
                    ))}
                    {initiatives.filter(i => i.status === status).length === 0 && <p style={{ fontSize: 12, color: C.textTertiary, textAlign: 'center', padding: 16 }}>None</p>}
                  </div>
                </div>
              ))}
            </div>

            <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
              <div style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.cardBorder}` }}>
                <h4 style={{ margin: '0 0 12px', fontSize: 13, fontWeight: 700, color: STATUS_COLORS['Done'] }}>Recently Completed</h4>
                {initiatives.filter(i => i.status === 'Done').length === 0 && <p style={{ fontSize: 12, color: C.textTertiary }}>None yet</p>}
                {initiatives.filter(i => i.status === 'Done').slice(0, 5).map(i => (
                  <div key={i.id} style={{ padding: '6px 0', borderBottom: `1px solid ${C.divider}`, fontSize: 12, color: C.textPrimary }}>
                    {i.title} <span style={{ color: C.textTertiary }}>({i.category})</span>
                  </div>
                ))}
              </div>
              <div style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.cardBorder}` }}>
                <h4 style={{ margin: '0 0 12px', fontSize: 13, fontWeight: 700, color: '#F59E0B' }}>Upcoming Due Dates</h4>
                {initiatives.filter(i => i.due_date && i.status !== 'Done').sort((a, b) => new Date(a.due_date) - new Date(b.due_date)).slice(0, 5).map(i => (
                  <div key={i.id} style={{ padding: '6px 0', borderBottom: `1px solid ${C.divider}`, fontSize: 12, display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ color: C.textPrimary }}>{i.title}</span>
                    <span style={{ color: new Date(i.due_date) < new Date() ? '#D81F26' : C.textTertiary, fontWeight: 600 }}>{new Date(i.due_date).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}</span>
                  </div>
                ))}
                {initiatives.filter(i => i.due_date && i.status !== 'Done').length === 0 && <p style={{ fontSize: 12, color: C.textTertiary }}>No upcoming dates</p>}
              </div>
            </div>

            {initiatives.filter(i => i.status === 'Blocked' || i.priority === 'Urgent').length > 0 && (
              <div style={{ background: '#FDE8E8', borderRadius: 6, padding: 16, border: '1px solid #FCA5A5' }}>
                <h4 style={{ margin: '0 0 12px', fontSize: 13, fontWeight: 700, color: '#D81F26' }}>Needs Attention</h4>
                {initiatives.filter(i => i.status === 'Blocked' || i.priority === 'Urgent').map(i => (
                  <div key={i.id} style={{ padding: '6px 0', borderBottom: '1px solid #FCA5A5', fontSize: 12, display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span style={{ color: C.textPrimary, fontWeight: 500 }}>{i.title}</span>
                    <span style={{ fontSize: 10, padding: '2px 6px', borderRadius: 4, background: i.status === 'Blocked' ? '#D81F26' : '#F59E0B', color: '#fff', fontWeight: 600 }}>{i.status === 'Blocked' ? 'BLOCKED' : 'URGENT'}</span>
                  </div>
                ))}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'center', paddingTop: 8 }}>
              <button onClick={() => setShowMetricPicker(true)} style={{ padding: '12px 28px', borderRadius: 6, border: 'none', background: C.primary, color: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8 }}>
                Generate Weekly Pitstop Presentation
              </button>
            </div>
          </div>
        );
      })() : (
      <div style={{ background: C.cardBg, borderRadius: 6, border: `1px solid ${C.cardBorder}`, overflow: 'hidden' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.divider }}>
              {['Status','Title','Priority','Owner','Due Date','Category',''].map(h => (
                <th key={h} style={{ padding: '10px 12px', textAlign: 'left', fontSize: 11, fontWeight: 600, color: C.textSecondary, borderBottom: `1px solid ${C.cardBorder}` }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} style={{ padding: 30, textAlign: 'center', color: C.textTertiary, fontSize: 13 }}>No initiatives yet. Click "+ Add Initiative" to get started.</td></tr>
            )}
            {filtered.map(item => (
              <React.Fragment key={item.id}>
                <tr onClick={() => toggleExpand(item.id)} style={{ borderBottom: `1px solid ${C.divider}`, cursor: 'pointer', background: expandedId === item.id ? '#F8FAFC' : 'transparent', transition: 'background 0.15s' }}>
                  <td style={{ padding: '10px 12px' }}>
                    <select value={item.status} onClick={e => e.stopPropagation()} onChange={e => updateStatus(item.id, e.target.value)} style={{ padding: '3px 8px', borderRadius: 6, border: 'none', background: STATUS_COLORS[item.status], color: '#fff', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>
                      {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: '10px 12px', fontWeight: 500, color: C.textPrimary }}>{item.title}</td>
                  <td style={{ padding: '10px 12px' }}>
                    <span style={{ padding: '2px 10px', borderRadius: 6, border: `1px solid ${PRIORITY_COLORS[item.priority]}`, color: PRIORITY_COLORS[item.priority], fontSize: 11, fontWeight: 600, background: item.priority === 'Urgent' ? '#FDE8E8' : 'transparent' }}>
                      {item.priority}
                    </span>
                  </td>
                  <td style={{ padding: '10px 12px', color: C.textSecondary, fontSize: 12 }}>{item.owner || '—'}</td>
                  <td style={{ padding: '10px 12px', color: C.textSecondary, fontSize: 12 }}>{item.due_date ? new Date(item.due_date).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '—'}</td>
                  <td style={{ padding: '10px 12px' }}>
                    <span style={{ padding: '2px 8px', borderRadius: 4, background: C.divider, color: C.textSecondary, fontSize: 11, fontWeight: 500 }}>{item.category}</span>
                  </td>
                  <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                    <button onClick={e => { e.stopPropagation(); setEditItem(item); setFormData({ title: item.title, description: item.description || '', status: item.status, priority: item.priority, owner: item.owner || '', due_date: item.due_date ? item.due_date.slice(0, 10) : '', category: item.category || 'General' }); setShowForm(true); }} style={{ background: 'none', border: 'none', color: C.textTertiary, cursor: 'pointer', fontSize: 12, marginRight: 6 }}>Edit</button>
                    <button onClick={e => { e.stopPropagation(); deleteInitiative(item.id); }} style={{ background: 'none', border: 'none', color: C.danger, cursor: 'pointer', fontSize: 12 }}>Delete</button>
                  </td>
                </tr>
                {expandedId === item.id && (
                  <tr><td colSpan={7} style={{ padding: '0 12px 16px', background: '#F8FAFC' }}>
                    {item.description && <p style={{ margin: '8px 0 12px', fontSize: 13, color: C.textSecondary, lineHeight: 1.5 }}>{item.description}</p>}
                    <div style={{ borderTop: `1px solid ${C.cardBorder}`, paddingTop: 12 }}>
                      <h5 style={{ margin: '0 0 8px', fontSize: 12, fontWeight: 600, color: C.textSecondary }}>Comments</h5>
                      {(comments[item.id] || []).map(c => (
                        <div key={c.id} style={{ display: 'flex', gap: 8, marginBottom: 8, padding: '8px 10px', background: C.cardBg, borderRadius: 4, border: `1px solid ${C.cardBorder}` }}>
                          <div style={{ width: 28, height: 28, borderRadius: '50%', background: C.primary, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{(c.author || '?')[0].toUpperCase()}</div>
                          <div>
                            <div style={{ display: 'flex', gap: 8, alignItems: 'baseline' }}>
                              <span style={{ fontWeight: 600, fontSize: 12, color: C.textPrimary }}>{c.author}</span>
                              <span style={{ fontSize: 10, color: C.textTertiary }}>{new Date(c.created_at).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</span>
                            </div>
                            <p style={{ margin: '2px 0 0', fontSize: 12, color: C.textSecondary, lineHeight: 1.4 }}>{c.content}</p>
                          </div>
                        </div>
                      ))}
                      <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                        <input value={commentText} onChange={e => setCommentText(e.target.value)} onKeyDown={e => e.key === 'Enter' && addComment(item.id)} placeholder="Add a comment..." style={{ flex: 1, padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, fontFamily: 'inherit', background: C.cardBg, color: C.textPrimary }} />
                        <button onClick={() => addComment(item.id)} disabled={!commentText.trim()} style={{ padding: '8px 16px', borderRadius: 4, border: 'none', background: commentText.trim() ? C.primary : C.textTertiary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: commentText.trim() ? 'pointer' : 'default' }}>Send</button>
                      </div>
                    </div>
                    <p style={{ margin: '8px 0 0', fontSize: 10, color: C.textTertiary }}>Created by {item.created_by} on {new Date(item.created_at).toLocaleDateString('en-GB')}</p>
                  </td></tr>
                )}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
      )}

      {/* ─── Metric Picker Modal ─── */}
      {showMetricPicker && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', zIndex: 190, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }} onClick={() => setShowMetricPicker(false)}>
          <div style={{ background: C.cardBg, borderRadius: 12, width: '100%', maxWidth: 560, maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }} onClick={e => e.stopPropagation()}>
            <div style={{ padding: '24px 24px 16px', borderBottom: `1px solid ${C.cardBorder}` }}>
              <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.textPrimary }}>Select Metrics for Presentation</h3>
              <p style={{ margin: '8px 0 0', fontSize: 12, color: C.textSecondary }}>Choose which dashboard metrics to include alongside your initiatives.</p>
            </div>

            {/* Date range picker */}
            <div style={{ margin: '16px 24px 0', padding: '14px 16px', borderRadius: 8, background: '#E8F0EB', border: '1px solid #B6D4C0' }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: '#124A2B', marginBottom: 10 }}>📅 Date Range</div>
              <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                <input type="date" value={dashboardState?.dateRange?.start || ''} onChange={e => dispatch && dispatch({ type: 'SET_DATE_RANGE', payload: { ...dashboardState.dateRange, start: e.target.value } })} style={{ padding: '6px 10px', fontSize: 12, border: '1px solid #B6D4C0', borderRadius: 4, background: '#fff', color: C.textPrimary, fontFamily: 'inherit' }} />
                <span style={{ fontSize: 12, color: C.textSecondary }}>to</span>
                <input type="date" value={dashboardState?.dateRange?.end || ''} onChange={e => dispatch && dispatch({ type: 'SET_DATE_RANGE', payload: { ...dashboardState.dateRange, end: e.target.value } })} style={{ padding: '6px 10px', fontSize: 12, border: '1px solid #B6D4C0', borderRadius: 4, background: '#fff', color: C.textPrimary, fontFamily: 'inherit' }} />
              </div>
              <div style={{ marginTop: 10 }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: '#124A2B', marginBottom: 6 }}>Comparison</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  {[{ val: 'none', label: 'None' }, { val: 'previous_period', label: 'Previous Period' }, { val: 'previous_month', label: 'Previous Month' }, { val: 'previous_year', label: 'Previous Year' }].map(opt => (
                    <button key={opt.val} onClick={() => dispatch && dispatch({ type: 'SET_COMPARISON', payload: opt.val })} style={{ padding: '4px 10px', fontSize: 11, fontWeight: (dashboardState?.comparison || 'none') === opt.val ? 700 : 500, border: `1px solid ${(dashboardState?.comparison || 'none') === opt.val ? C.primary : '#B6D4C0'}`, borderRadius: 4, background: (dashboardState?.comparison || 'none') === opt.val ? C.primary : 'transparent', color: (dashboardState?.comparison || 'none') === opt.val ? '#fff' : C.textSecondary, cursor: 'pointer' }}>{opt.label}</button>
                  ))}
                </div>
              </div>
            </div>

            {/* Select All / Clear All */}
            <div style={{ display: 'flex', gap: 8, padding: '12px 24px 0' }}>
              <button onClick={() => setSelectedMetricGroups(prev => { const n = { ...prev }; METRIC_GROUPS.forEach(g => n[g.key] = true); return n; })} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Select All</button>
              <button onClick={() => setSelectedMetricGroups(prev => { const n = { ...prev }; METRIC_GROUPS.forEach(g => n[g.key] = false); return n; })} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Clear All</button>
            </div>

            {/* Metric group checklist */}
            <div style={{ padding: '12px 24px', display: 'flex', flexDirection: 'column', gap: 8 }}>
              {METRIC_GROUPS.map(group => (
                <div key={group.key} onClick={() => setSelectedMetricGroups(prev => ({ ...prev, [group.key]: !prev[group.key] }))} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '14px 16px', borderRadius: 8, border: `2px solid ${selectedMetricGroups[group.key] ? C.primary : C.cardBorder}`, background: selectedMetricGroups[group.key] ? '#E8F0EB' : C.cardBg, cursor: 'pointer', transition: 'all 0.15s' }}>
                  <div style={{ width: 22, height: 22, borderRadius: 4, border: `2px solid ${selectedMetricGroups[group.key] ? C.primary : C.textTertiary}`, background: selectedMetricGroups[group.key] ? C.primary : 'transparent', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, transition: 'all 0.15s' }}>
                    {selectedMetricGroups[group.key] && <span style={{ color: '#fff', fontSize: 14, fontWeight: 700 }}>✓</span>}
                  </div>
                  <div style={{ fontSize: 20 }}>{group.icon}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 14, fontWeight: 600, color: C.textPrimary }}>{group.label}</div>
                    <div style={{ fontSize: 11, color: C.textTertiary, marginTop: 2 }}>{group.metrics.map(m => m.label).slice(0, 3).join(', ')}{group.metrics.length > 3 ? ` +${group.metrics.length - 3} more` : ''}</div>
                  </div>
                </div>
              ))}

              {/* Initiatives row - always included */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '14px 16px', borderRadius: 8, border: `2px solid ${C.primary}`, background: '#E8F0EB', opacity: 0.7 }}>
                <div style={{ width: 22, height: 22, borderRadius: 4, border: `2px solid ${C.primary}`, background: C.primary, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                  <span style={{ color: '#fff', fontSize: 14, fontWeight: 700 }}>✓</span>
                </div>
                <div style={{ fontSize: 20 }}>📋</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, color: C.textPrimary }}>Initiatives</div>
                  <div style={{ fontSize: 11, color: C.textTertiary, marginTop: 2 }}>Always included — summary, categories, activity highlights</div>
                </div>
              </div>
            </div>

            {/* Action buttons */}
            <div style={{ padding: '16px 24px 24px', borderTop: `1px solid ${C.cardBorder}`, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
              <button onClick={() => setShowMetricPicker(false)} style={{ padding: '10px 24px', borderRadius: 6, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 14, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
              <button onClick={generatePresentationSlides} style={{ padding: '10px 28px', borderRadius: 6, border: 'none', background: C.primary, color: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}>Generate Presentation</button>
            </div>
          </div>
        </div>
      )}

      {/* ─── Presentation Editor Modal ─── */}
      {showPresentationEditor && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.7)', zIndex: 200, display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={() => setShowPresentationEditor(false)}>
          <div className="crm-pres-editor" style={{ display: 'flex', width: '95vw', maxWidth: 1200, height: '90vh', background: C.cardBg, borderRadius: 10, overflow: 'hidden', boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }} onClick={e => e.stopPropagation()}>

            {/* Left sidebar: slide list */}
            <div className="crm-pres-sidebar" style={{ width: 260, borderRight: `1px solid ${C.cardBorder}`, background: C.pageBg, display: 'flex', flexDirection: 'column', overflow: 'hidden', flexShrink: 0 }}>
              <div style={{ padding: 16, borderBottom: `1px solid ${C.cardBorder}` }}>
                <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Slides ({presentationSlides.length})</h3>
              </div>
              <div style={{ flex: 1, overflowY: 'auto', padding: 8, display: 'flex', flexDirection: 'column', gap: 6 }}>
                {presentationSlides.map((slide, idx) => (
                  <div key={slide.id} onClick={() => setActiveSlideIndex(idx)} style={{ padding: 10, borderRadius: 6, border: `2px solid ${idx === activeSlideIndex ? C.primary : C.cardBorder}`, background: idx === activeSlideIndex ? '#E8F0EB' : C.cardBg, cursor: 'pointer', position: 'relative' }}>
                    <div style={{ fontSize: 11, fontWeight: 600, color: C.textPrimary, marginBottom: 4, paddingRight: 50 }}>{idx + 1}. {slide.title || 'Untitled'}</div>
                    <div style={{ fontSize: 10, color: C.textTertiary, textTransform: 'capitalize' }}>{slide.type}</div>
                    <div style={{ position: 'absolute', top: 6, right: 6, display: 'flex', gap: 2 }}>
                      {idx > 0 && <button onClick={e => { e.stopPropagation(); moveSlide(idx, idx - 1); }} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: C.textTertiary, padding: '2px 4px' }}>{'\u25B2'}</button>}
                      {idx < presentationSlides.length - 1 && <button onClick={e => { e.stopPropagation(); moveSlide(idx, idx + 1); }} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: C.textTertiary, padding: '2px 4px' }}>{'\u25BC'}</button>}
                      <button onClick={e => { e.stopPropagation(); deleteSlide(idx); }} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: C.danger, padding: '2px 4px' }}>{'\u2715'}</button>
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ padding: 12, borderTop: `1px solid ${C.cardBorder}`, display: 'flex', gap: 6 }}>
                <button onClick={addNewSlide} style={{ flex: 1, padding: '8px 0', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>+ Add Slide</button>
              </div>
            </div>

            {/* Right panel: slide editor */}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ padding: 16, borderBottom: `1px solid ${C.cardBorder}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: C.textPrimary }}>Edit Slide</h3>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button onClick={downloadPptx} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>{'\u2B07'} Download PPTX</button>
                  <button onClick={() => setShowPresentationEditor(false)} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Close</button>
                </div>
              </div>

              {presentationSlides.length > 0 && activeSlideIndex < presentationSlides.length && (() => {
                const slide = presentationSlides[activeSlideIndex];
                return (
                  <div className="crm-pres-editor-panel" style={{ flex: 1, overflowY: 'auto', padding: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
                    {/* Title */}
                    <div>
                      <label style={{ fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4, display: 'block' }}>Slide Title</label>
                      <input value={slide.title} onChange={e => updateSlide(activeSlideIndex, { title: e.target.value })} style={{ width: '100%', padding: '10px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 15, fontWeight: 600, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' }} />
                    </div>
                    {/* Content */}
                    <div>
                      <label style={{ fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4, display: 'block' }}>Content</label>
                      <textarea value={slide.content} onChange={e => updateSlide(activeSlideIndex, { content: e.target.value })} rows={4} style={{ width: '100%', padding: '10px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary, resize: 'vertical', boxSizing: 'border-box' }} />
                    </div>
                    {/* Bullets */}
                    <div>
                      <label style={{ fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4, display: 'block' }}>Bullet Points</label>
                      {slide.bullets.map((bullet, bIdx) => (
                        <div key={bIdx} style={{ display: 'flex', gap: 6, marginBottom: 4, alignItems: 'center' }}>
                          <span style={{ color: C.textTertiary, fontSize: 14 }}>{'\u2022'}</span>
                          <input value={bullet} onChange={e => { const nb = [...slide.bullets]; nb[bIdx] = e.target.value; updateSlide(activeSlideIndex, { bullets: nb }); }} style={{ flex: 1, padding: '6px 10px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary }} />
                          <button onClick={() => { updateSlide(activeSlideIndex, { bullets: slide.bullets.filter((_, i) => i !== bIdx) }); }} style={{ background: 'none', border: 'none', color: C.danger, cursor: 'pointer', fontSize: 16, padding: '2px 6px' }}>{'\u00D7'}</button>
                        </div>
                      ))}
                      <button onClick={() => updateSlide(activeSlideIndex, { bullets: [...slide.bullets, ''] })} style={{ fontSize: 12, color: C.primary, background: 'none', border: 'none', cursor: 'pointer', fontWeight: 600, marginTop: 4 }}>+ Add Bullet</button>
                    </div>
                    {/* Image upload */}
                    <div>
                      <label style={{ fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4, display: 'block' }}>Screenshot / Image</label>
                      {slide.image ? (
                        <div style={{ position: 'relative', display: 'inline-block' }}>
                          <img src={slide.image.dataUrl} alt={slide.image.name} style={{ maxWidth: '100%', maxHeight: 300, borderRadius: 6, border: `1px solid ${C.cardBorder}` }} />
                          <button onClick={() => updateSlide(activeSlideIndex, { image: null })} style={{ position: 'absolute', top: 8, right: 8, background: 'rgba(0,0,0,0.6)', color: '#fff', border: 'none', borderRadius: '50%', width: 26, height: 26, cursor: 'pointer', fontSize: 14, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{'\u00D7'}</button>
                          <p style={{ fontSize: 11, color: C.textTertiary, marginTop: 4 }}>{slide.image.name}</p>
                        </div>
                      ) : (
                        <div onClick={() => slideImageRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) { const dt = new DataTransfer(); dt.items.add(f); slideImageRef.current.files = dt.files; handleSlideImageUpload({ target: { files: [f] } }); } }} style={{ border: `2px dashed ${C.cardBorder}`, borderRadius: 6, padding: 30, textAlign: 'center', cursor: 'pointer', color: C.textTertiary, fontSize: 13, transition: 'border-color 0.2s' }}>
                          Click or drag to upload a screenshot (PNG, JPEG, GIF, WebP - max 20MB)
                        </div>
                      )}
                      <input ref={slideImageRef} type="file" accept="image/png,image/jpeg,image/gif,image/webp" style={{ display: 'none' }} onChange={handleSlideImageUpload} />
                    </div>
                    {/* Preview */}
                    <div style={{ background: C.divider, borderRadius: 8, padding: 20, border: `1px solid ${C.cardBorder}` }}>
                      <div style={{ fontSize: 11, fontWeight: 600, color: C.textTertiary, marginBottom: 8, textTransform: 'uppercase', letterSpacing: 1 }}>Preview</div>
                      <div style={{ background: '#fff', borderRadius: 6, padding: 24, minHeight: 180, boxShadow: '0 2px 8px rgba(0,0,0,0.08)', position: 'relative', overflow: 'hidden' }}>
                        {/* Title slide visual */}
                        {slide.type === 'title' && slide.titleData ? (
                          <div>
                            <div style={{ background: C.primary, borderRadius: '6px 6px 0 0', padding: '32px 24px 24px', margin: '-24px -24px 0 -24px' }}>
                              <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#fff' }}>{slide.title || 'Untitled'}</h2>
                              <div style={{ fontSize: 14, color: 'rgba(255,255,255,0.8)', marginTop: 10 }}>{slide.titleData.dateRange}</div>
                              {slide.titleData.compLabel && <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.6)', marginTop: 4, fontStyle: 'italic' }}>{slide.titleData.compLabel}</div>}
                            </div>
                            <div style={{ padding: '16px 0 0' }}>
                              <div style={{ width: 40, height: 2, background: C.primary, marginBottom: 10 }} />
                              <div style={{ fontSize: 16, fontWeight: 700, color: C.primary }}>omni.pet</div>
                            </div>
                          </div>
                        ) : slide.type === 'metrics' && slide.metricData && slide.metricData.length > 0 ? (
                          /* KPI Card Grid preview */
                          <div>
                            <div style={{ background: C.primary, borderRadius: '6px 6px 0 0', padding: '10px 20px', margin: '-24px -24px 16px -24px' }}>
                              <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#fff' }}>{slide.title || 'Metrics'}</h2>
                            </div>
                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(140px, 1fr))', gap: 8 }}>
                              {slide.metricData.map((m, i) => (
                                <div key={i} style={{ background: '#F8F9FA', borderRadius: 6, padding: '12px 14px', border: '1px solid #E5E5EB' }}>
                                  <div style={{ fontSize: 18, fontWeight: 700, color: '#272D45', marginBottom: 4, lineHeight: 1.2 }}>{m.formatted}</div>
                                  <div style={{ fontSize: 10, color: '#676986', marginBottom: 6 }}>{m.label}</div>
                                  {m.change !== null && m.change !== undefined && (
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                      <span style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, fontSize: 9, fontWeight: 600, color: '#fff', background: m.arrow === '\u2191' ? '#18917B' : m.arrow === '\u2193' ? '#D81F26' : '#94A3B8' }}>
                                        {m.arrow} {Math.abs(m.change).toFixed(1)}%
                                      </span>
                                      {m.prevFormatted && <span style={{ fontSize: 8, color: '#9CA3AF' }}>vs {m.prevFormatted}</span>}
                                    </div>
                                  )}
                                </div>
                              ))}
                            </div>
                          </div>
                        ) : slide.type === 'summary' && slide.initiativeData ? (
                          /* Initiatives Summary visual */
                          <div>
                            <div style={{ background: C.primary, borderRadius: '6px 6px 0 0', padding: '10px 20px', margin: '-24px -24px 16px -24px' }}>
                              <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#fff' }}>{slide.title || 'Summary'}</h2>
                            </div>
                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))', gap: 8, marginBottom: 12 }}>
                              {[
                                { label: 'Total', value: slide.initiativeData.total, color: '#272D45' },
                                { label: 'Done', value: slide.initiativeData.done, color: '#124A2B' },
                                { label: 'In Progress', value: slide.initiativeData.inProgress, color: '#18917B' },
                                { label: 'Blocked', value: slide.initiativeData.blocked, color: '#D81F26' },
                              ].map((c, i) => (
                                <div key={i} style={{ background: '#F8F9FA', borderRadius: 6, padding: 10, border: '1px solid #E5E5EB', textAlign: 'center' }}>
                                  <div style={{ fontSize: 22, fontWeight: 700, color: c.color }}>{c.value}</div>
                                  <div style={{ fontSize: 9, color: '#676986', marginTop: 2 }}>{c.label}</div>
                                </div>
                              ))}
                            </div>
                            <div style={{ fontSize: 10, fontWeight: 600, color: '#272D45', marginBottom: 4 }}>Completion: {slide.initiativeData.completionRate.toFixed(0)}%</div>
                            <div style={{ background: '#E5E7EB', borderRadius: 4, height: 8, overflow: 'hidden' }}>
                              <div style={{ width: `${slide.initiativeData.completionRate}%`, height: '100%', background: '#124A2B', borderRadius: 4, transition: 'width 0.3s' }} />
                            </div>
                            {slide.initiativeData.statusDistribution.length > 0 && (
                              <div style={{ display: 'flex', gap: 12, marginTop: 10, flexWrap: 'wrap' }}>
                                {slide.initiativeData.statusDistribution.map((s, i) => (
                                  <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 10, color: '#676986' }}>
                                    <div style={{ width: 8, height: 8, borderRadius: '50%', background: `#${s.color}` }} />
                                    {s.name}: {s.value}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        ) : (
                          /* Default bullet list preview */
                          <div>
                            <div style={{ background: C.primary, borderRadius: '6px 6px 0 0', padding: '10px 20px', margin: '-24px -24px 16px -24px' }}>
                              <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#fff' }}>{slide.title || 'Untitled'}</h2>
                            </div>
                            {slide.content && <p style={{ margin: '0 0 12px', fontSize: 13, color: C.textSecondary, whiteSpace: 'pre-line' }}>{slide.content}</p>}
                            {slide.bullets.length > 0 && (
                              <ul style={{ margin: '8px 0', paddingLeft: 20 }}>
                                {slide.bullets.filter(b => b.trim()).map((b, i) => <li key={i} style={{ fontSize: 12, color: C.textPrimary, marginBottom: 4 }}>{b}</li>)}
                              </ul>
                            )}
                          </div>
                        )}
                        {slide.image && <img src={slide.image.dataUrl} alt="" style={{ maxWidth: '100%', maxHeight: 200, borderRadius: 4, marginTop: 8 }} />}
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── SECTION COMPONENTS ───

function OverviewSection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const { start, end } = state.dateRange;
  const [showCostForm, setShowCostForm] = useState(false);
  const [costForm, setCostForm] = useState({ month: '', channel: 'Emails', cost: '', notes: '' });
  const [costError, setCostError] = useState(null);
  const compRange = useMemo(() => computeComparisonRange(start, end, state.comparison), [start, end, state.comparison]);
  const compLabel = COMP_LABELS[state.comparison] || null;
  const fEmailFlows = useMemo(() => filterByDateRange(state.emailFlows, start, end, 'week'), [state.emailFlows, start, end]);
  const fOutreach = useMemo(() => filterByDateRange(state.outreach, start, end, 'week'), [state.outreach, start, end]);
  const fRevenue = useMemo(() => filterByDateRange(state.revenue, start, end, 'week'), [state.revenue, start, end]);
  const fLoyalty = useMemo(() => filterByDateRange(state.loyalty, start, end, 'month'), [state.loyalty, start, end]);
  const fSegments = useMemo(() => filterByDateRange(state.segments, start, end, 'month'), [state.segments, start, end]);

  // Comparison period data
  const cEmailFlows = useMemo(() => compRange ? filterByDateRange(state.emailFlows, compRange.start, compRange.end, 'week') : [], [state.emailFlows, compRange]);
  const cOutreach = useMemo(() => compRange ? filterByDateRange(state.outreach, compRange.start, compRange.end, 'week') : [], [state.outreach, compRange]);
  const cRevenue = useMemo(() => compRange ? filterByDateRange(state.revenue, compRange.start, compRange.end, 'week') : [], [state.revenue, compRange]);
  const cLoyalty = useMemo(() => compRange ? filterByDateRange(state.loyalty, compRange.start, compRange.end, 'month') : [], [state.loyalty, compRange]);
  const cSegments = useMemo(() => compRange ? filterByDateRange(state.segments, compRange.start, compRange.end, 'month') : [], [state.segments, compRange]);

  const latestWeek = [...new Set(fEmailFlows.map(r => r.week))].sort().pop();
  const latestEmailRev = fEmailFlows.filter(r => r.week === latestWeek).reduce((s, r) => s + (r.revenue || 0), 0);
  const latestOutreachWeek = [...new Set(fOutreach.map(r => r.week))].sort().pop();
  const latestOutreachRev = fOutreach.filter(r => r.week === latestOutreachWeek).reduce((s, r) => s + (r.revenue || 0), 0);
  const crmRevWeek = latestEmailRev + latestOutreachRev;
  const latestTotalRev = fRevenue.length > 0 ? fRevenue[fRevenue.length - 1].totalRevenue : 1;
  const crmPct = (crmRevWeek / latestTotalRev) * 100;
  const latestLoyalty = fLoyalty.length > 0 ? fLoyalty[fLoyalty.length - 1] : null;
  const latestSeg = fSegments.length > 0 ? fSegments[fSegments.length - 1] : null;
  const totalCost = state.activityROI.reduce((s, r) => s + r.totalCost, 0);
  const totalIncRev = state.activityROI.reduce((s, r) => s + r.incrementalRevenue, 0);
  const avgROI = totalCost > 0 ? totalIncRev / totalCost : 0;
  const activeTests = state.holdoutTests.filter(t => t.status === 'active').length;

  // Comparison KPI values
  const comp = useMemo(() => {
    if (!compRange) return null;
    const cLatestWeek = [...new Set(cEmailFlows.map(r => r.week))].sort().pop();
    const cLatestEmailRev = cEmailFlows.filter(r => r.week === cLatestWeek).reduce((s, r) => s + (r.revenue || 0), 0);
    const cLatestOutWeek = [...new Set(cOutreach.map(r => r.week))].sort().pop();
    const cLatestOutRev = cOutreach.filter(r => r.week === cLatestOutWeek).reduce((s, r) => s + (r.revenue || 0), 0);
    const cCrmRev = cLatestEmailRev + cLatestOutRev;
    const cTotalRev = cRevenue.length > 0 ? cRevenue[cRevenue.length - 1].totalRevenue : 1;
    const cCrmPct = (cCrmRev / cTotalRev) * 100;
    const cLatestLoy = cLoyalty.length > 0 ? cLoyalty[cLoyalty.length - 1] : null;
    const cLatestSeg = cSegments.length > 0 ? cSegments[cSegments.length - 1] : null;
    return { crmRev: cCrmRev, crmPct: cCrmPct, members: cLatestLoy?.totalMembers, atRisk: cLatestSeg?.segAtRisk };
  }, [compRange, cEmailFlows, cOutreach, cRevenue, cLoyalty, cSegments]);

  const allWeeks = [...new Set([...fEmailFlows.map(r => r.week), ...fOutreach.map(r => r.week)])].sort();
  // Build comparison CRM revenue trend (indexed by position, not date)
  const compCrmWeeks = compRange ? [...new Set([...cEmailFlows.map(r => r.week), ...cOutreach.map(r => r.week)])].sort() : [];
  const compCrmValues = compCrmWeeks.map(w => {
    const eRev = cEmailFlows.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0);
    const oRev = cOutreach.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0);
    return eRev + oRev;
  });
  const crmTrend = allWeeks.map((w, i) => {
    const emailRev = fEmailFlows.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0);
    const smsRev = fOutreach.filter(r => r.week === w && r.channel === 'SMS Blast').reduce((s, r) => s + (r.revenue || 0), 0);
    const whatsappRev = fOutreach.filter(r => r.week === w && r.channel === 'WhatsApp').reduce((s, r) => s + (r.revenue || 0), 0);
    const postcardRev = fOutreach.filter(r => r.week === w && r.channel === 'Postcard').reduce((s, r) => s + (r.revenue || 0), 0);
    const totalRev = fRevenue.find(r => r.week === w)?.totalRevenue || 0;
    const entry = { week: w.slice(5), emailRevenue: emailRev, smsRevenue: smsRev, whatsappRevenue: whatsappRev, postcardRevenue: postcardRev, totalRevenue: totalRev };
    if (compCrmValues.length > 0) entry.prevCrmRevenue = compCrmValues[i] ?? null;
    return entry;
  });

  const roiSorted = [...state.activityROI].sort((a, b) => b.incrementalRevenue - a.incrementalRevenue);
  const alerts = useMemo(() => generateAlerts(state), [state]);
  const roiColor = (v) => v > 10 ? '#18917B' : v > 3 ? '#2D8B6E' : v > 1 ? '#F59E0B' : '#D81F26';

  // ─── Channel Cost & ROI (monthly) ───
  // Build monthly revenue per channel from weekly data
  const monthlyChannelData = useMemo(() => {
    const costs = state.channelCosts || [];
    // Group weekly revenue into months per channel
    const monthRevMap = {};
    const addRev = (week, channel, rev) => {
      if (!week || !rev) return;
      const m = week.slice(0, 7); // "2025-01-06" → "2025-01"
      if (!monthRevMap[m]) monthRevMap[m] = { Emails: 0, SMS: 0, WhatsApp: 0, Postcards: 0 };
      monthRevMap[m][channel] += rev;
    };
    fEmailFlows.forEach(r => addRev(r.week, 'Emails', r.revenue || 0));
    fOutreach.filter(r => r.channel === 'SMS Blast').forEach(r => addRev(r.week, 'SMS', r.revenue || 0));
    fOutreach.filter(r => r.channel === 'WhatsApp').forEach(r => addRev(r.week, 'WhatsApp', r.revenue || 0));
    fOutreach.filter(r => r.channel === 'Postcard').forEach(r => addRev(r.week, 'Postcards', r.revenue || 0));
    // All months from costs + revenue
    const allMonths = [...new Set([...costs.map(c => c.month), ...Object.keys(monthRevMap)])].sort();
    return allMonths.map(m => {
      const rev = monthRevMap[m] || { Emails: 0, SMS: 0, WhatsApp: 0, Postcards: 0 };
      const channels = CHANNEL_DEFS.map(ch => {
        const cost = costs.filter(c => c.month === m && c.channel === ch.key).reduce((s, c) => s + (Number(c.cost) || 0), 0);
        const revenue = rev[ch.key] || 0;
        const roi = cost > 0 ? revenue / cost : 0;
        return { ...ch, cost, revenue, roi };
      });
      return { month: m, channels };
    });
  }, [state.channelCosts, fEmailFlows, fOutreach]);
  const costMonths = monthlyChannelData.map(d => d.month);
  const [selectedCostMonth, setSelectedCostMonth] = useState(null);
  const activeCostMonth = selectedCostMonth || (costMonths.length > 0 ? costMonths[costMonths.length - 1] : null);
  const activeCostData = monthlyChannelData.find(d => d.month === activeCostMonth);
  const channelRoiChart = activeCostData ? activeCostData.channels.map(ch => ({ name: ch.key, revenue: ch.revenue, cost: ch.cost })) : [];

  const handleAddCost = () => {
    if (!costForm.month || !costForm.cost) { setCostError('Month and cost are required'); return; }
    dispatch({ type: 'APPEND_DATA', source: 'channelCosts', payload: [{ month: costForm.month, channel: costForm.channel, cost: Number(costForm.cost), notes: costForm.notes }] });
    setCostForm({ month: '', channel: 'Emails', cost: '', notes: '' });
    setCostError(null);
    setShowCostForm(false);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label="CRM Revenue" value={crmRevWeek} format="currency" status="good" compDelta={comp ? pctChange(crmRevWeek, comp.crmRev) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="CRM % of Revenue" value={crmPct} format="percent" status={crmPct < 20 ? 'warning' : 'good'} compDelta={comp ? pctChange(crmPct, comp.crmPct) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Active Milestone Members" value={latestLoyalty?.totalMembers} format="number" status="good" sparkData={fLoyalty} sparkKey="totalMembers" compDelta={comp ? pctChange(latestLoyalty?.totalMembers, comp.members) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="At-Risk Customers" value={latestSeg?.segAtRisk} format="number" status={latestSeg?.segAtRisk > 1200 ? 'warning' : 'good'} sparkData={fSegments} sparkKey="segAtRisk" compDelta={comp ? pctChange(latestSeg?.segAtRisk, comp.atRisk) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Avg CRM ROI" value={avgROI} format="multiplier" status={avgROI < 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label="Holdout Tests Running" value={activeTests} format="number" status="good" presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="CRM Revenue Contribution Over Time" tooltip="Weekly revenue by CRM channel (Emails, SMS, WhatsApp, Postcards) vs total revenue." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={crmTrend}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Area type="monotone" dataKey="emailRevenue" stackId="crm" fill="#124A2B" stroke="#124A2B" fillOpacity={0.6} name="Emails" />
            <Area type="monotone" dataKey="smsRevenue" stackId="crm" fill="#3B82F6" stroke="#3B82F6" fillOpacity={0.6} name="SMS" />
            <Area type="monotone" dataKey="whatsappRevenue" stackId="crm" fill="#25D366" stroke="#25D366" fillOpacity={0.6} name="WhatsApp" />
            <Area type="monotone" dataKey="postcardRevenue" stackId="crm" fill="#F59E0B" stroke="#F59E0B" fillOpacity={0.6} name="Postcards" />
            <Line type="monotone" dataKey="totalRevenue" stroke={C.textTertiary} strokeDasharray="5 5" strokeWidth={2} dot={false} name="Total Business Revenue" />
            {compRange && <Line type="monotone" dataKey="prevCrmRevenue" stroke={C.warning} strokeDasharray="6 3" strokeWidth={2} dot={false} name={`CRM Rev (${compLabel})`} connectNulls />}
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      {/* ─── Channel Cost & ROI (monthly) ─── */}
      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10, marginBottom: 16 }}>
          <ChartHeader title="Channel Cost & ROI" tooltip="Monthly costs per channel with computed ROI. Select a month to view that month's revenue vs cost." />
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {costMonths.length > 0 && (
              <select value={activeCostMonth || ''} onChange={e => setSelectedCostMonth(e.target.value)} style={{ padding: '6px 10px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 6, background: '#fff', color: C.textPrimary }}>
                {costMonths.map(m => <option key={m} value={m}>{new Date(m + '-01').toLocaleDateString('en-GB', { month: 'short', year: 'numeric' })}</option>)}
              </select>
            )}
            <button onClick={() => setShowCostForm(!showCostForm)} style={{ padding: '6px 14px', fontSize: 12, fontWeight: 600, border: `1px solid ${C.primary}`, background: showCostForm ? C.primary : 'transparent', color: showCostForm ? '#fff' : C.primary, borderRadius: 6, cursor: 'pointer' }}>
              {showCostForm ? 'Cancel' : '+ Add Cost'}
            </button>
          </div>
        </div>

        {showCostForm && (
          <div style={{ background: '#F9F7F2', borderRadius: 6, padding: 16, marginBottom: 16, border: `1px solid ${C.divider}` }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 12 }}>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: C.textSecondary, display: 'block', marginBottom: 4 }}>Month</label>
                <input type="month" value={costForm.month} onChange={e => setCostForm({ ...costForm, month: e.target.value })} style={{ width: '100%', padding: '6px 10px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', boxSizing: 'border-box' }} />
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: C.textSecondary, display: 'block', marginBottom: 4 }}>Channel</label>
                <select value={costForm.channel} onChange={e => setCostForm({ ...costForm, channel: e.target.value })} style={{ width: '100%', padding: '6px 10px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', boxSizing: 'border-box' }}>
                  {CHANNEL_DEFS.map(ch => <option key={ch.key} value={ch.key}>{ch.key}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: C.textSecondary, display: 'block', marginBottom: 4 }}>Cost (£)</label>
                <input type="number" value={costForm.cost} onChange={e => setCostForm({ ...costForm, cost: e.target.value })} placeholder="0" style={{ width: '100%', padding: '6px 10px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', boxSizing: 'border-box' }} />
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: C.textSecondary, display: 'block', marginBottom: 4 }}>Notes</label>
                <input type="text" value={costForm.notes} onChange={e => setCostForm({ ...costForm, notes: e.target.value })} placeholder="Optional" style={{ width: '100%', padding: '6px 10px', fontSize: 12, border: `1px solid ${C.cardBorder}`, borderRadius: 4, background: '#fff', boxSizing: 'border-box' }} />
              </div>
            </div>
            {costError && <p style={{ color: C.danger, fontSize: 11, margin: '8px 0 0' }}>{costError}</p>}
            <button onClick={handleAddCost} style={{ marginTop: 12, padding: '6px 20px', fontSize: 12, fontWeight: 600, background: C.primary, color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer' }}>Save Cost</button>
          </div>
        )}

        {activeCostData && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 12, marginBottom: 16 }}>
            {activeCostData.channels.map(ch => (
              <div key={ch.key} style={{ background: '#F9F7F2', borderRadius: 6, padding: 14, border: `1px solid ${C.divider}`, borderLeft: `4px solid ${ch.color}` }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.textPrimary, marginBottom: 8 }}>{ch.key}</div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                  <div>
                    <div style={{ fontSize: 10, color: C.textTertiary, textTransform: 'uppercase', letterSpacing: 0.5 }}>Revenue</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>{formatCurrency(ch.revenue)}</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: C.textTertiary, textTransform: 'uppercase', letterSpacing: 0.5 }}>Cost</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>{formatCurrency(ch.cost)}</div>
                  </div>
                  <div style={{ gridColumn: '1 / -1' }}>
                    <div style={{ fontSize: 10, color: C.textTertiary, textTransform: 'uppercase', letterSpacing: 0.5 }}>ROI</div>
                    <div style={{ fontSize: 18, fontWeight: 800, color: ch.roi > 5 ? '#18917B' : ch.roi > 2 ? '#F59E0B' : '#D81F26' }}>{ch.cost > 0 ? `${ch.roi.toFixed(1)}x` : 'N/A'}</div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={channelRoiChart} layout="vertical">
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <YAxis dataKey="name" type="category" width={80} tick={{ fontSize: 11, fill: C.textTertiary }} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="revenue" fill={C.primary} name="Revenue" radius={[0, 4, 4, 0]} />
            <Bar dataKey="cost" fill={C.danger} name="Cost" radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid crm-heatmap-grid" style={{ display: 'grid', gap: 16 }}>
        <div className="crm-card" style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Activity Performance Heatmap" tooltip="Visual heatmap of all CRM activities, colour-coded by performance level." />
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr>{['Activity','Channel','Cost','Attributed Rev','Incremental Rev','ROI','Customers'].map(h => (
                  <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Activity' || h === 'Channel' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {roiSorted.map((r, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                    <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{r.activity}</td>
                    <td style={{ padding: '8px 10px', color: C.textSecondary }}>{r.channel}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.totalCost)}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.attributedRevenue)}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(r.incrementalRevenue)}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                      <span style={{ background: roiColor(r.incrementalROI), color: '#fff', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>{formatMultiplier(r.incrementalROI)}</span>
                    </td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.customersInfluenced)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="CRM Alerts" tooltip="Auto-generated alerts based on threshold rules: email deliverability, at-risk segment growth, redemption rates, points liability, and holdout test confidence." />
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {alerts.map((a, i) => (
              <div key={i} style={{ padding: '10px 12px', borderRadius: 4, background: a.severity === 'danger' ? '#FDE8E8' : a.severity === 'warning' ? '#FFFBEB' : '#E8F5F0', borderLeft: `4px solid ${a.severity === 'danger' ? C.danger : a.severity === 'warning' ? C.warning : C.success}` }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontWeight: 600, fontSize: 12, color: C.textPrimary }}>{a.metric}</span>
                  {a.value && <span style={{ fontSize: 12, fontWeight: 600, color: a.severity === 'danger' ? C.danger : a.severity === 'warning' ? C.warning : C.success }}>{a.value}</span>}
                </div>
                <p style={{ margin: '4px 0 0', fontSize: 11, color: C.textSecondary }}>{a.message}</p>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function EmailFlowsSection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const [showAddFlow, setShowAddFlow] = useState(false);
  const [addFlowMode, setAddFlowMode] = useState('manual');
  const [flowForm, setFlowForm] = useState({ week: '', type: 'Flow', flowName: '', sends: '', openRate: '', ctr: '', unsubRate: '', revenue: '' });
  const [flowError, setFlowError] = useState(null);
  const [flowCsvPreview, setFlowCsvPreview] = useState(null);
  const [flowImageData, setFlowImageData] = useState(null);
  const [flowImagePreview, setFlowImagePreview] = useState(null);
  const [flowAiLoading, setFlowAiLoading] = useState(false);
  const [flowAiPreview, setFlowAiPreview] = useState(null);
  const flowImageRef = useRef(null);
  const flowCsvRef = useRef(null);

  const submitManualFlow = () => {
    if (!flowForm.week || !flowForm.type) { setFlowError('Week and type are required'); return; }
    if (flowForm.type === 'Flow' && !flowForm.flowName) { setFlowError('Flow name required'); return; }
    const row = { ...flowForm };
    ['sends', 'openRate', 'ctr', 'unsubRate', 'revenue'].forEach(f => { row[f] = row[f] ? Number(row[f]) : 0; });
    row.delivered = row.sends; row.opens = Math.round(row.sends * row.openRate / 100);
    row.clicks = Math.round(row.sends * row.ctr / 100); row.unsubscribes = Math.round(row.sends * row.unsubRate / 100);
    row.conversions = 0; row.listSize = 0;
    dispatch({ type: 'APPEND_DATA', source: 'emailFlows', payload: [row] });
    setShowAddFlow(false); setFlowForm({ week: '', type: 'Flow', flowName: '', sends: '', openRate: '', ctr: '', unsubRate: '', revenue: '' }); setFlowError(null);
  };

  const handleFlowImage = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!['image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(file.type)) { setFlowError('Use PNG, JPEG, GIF, or WebP'); return; }
    if (file.size > 20 * 1024 * 1024) { setFlowError('Max 20MB'); return; }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const base64 = dataUrl.split(',')[1];
      setFlowImageData({ base64, mediaType: file.type, name: file.name });
      setFlowImagePreview(dataUrl);
    };
    reader.readAsDataURL(file);
  };

  const organizeFlowScreenshot = async () => {
    const apiKey = getAnthropicKey();
    if (!apiKey) { setFlowError('No API key configured. Open Settings.'); return; }
    if (!flowImageData) { setFlowError('Upload a screenshot first'); return; }
    setFlowAiLoading(true); setFlowError(null);
    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST', headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514', max_tokens: 4096,
          system: 'Extract email flow/campaign data from screenshots. Return ONLY a JSON array of objects. Each object must have: week (YYYY-MM-DD), type (Flow or Campaign), flowName (string, empty for Campaign), sends (number), openRate (number), ctr (number), unsubRate (number), revenue (number). No markdown, no explanation.',
          messages: [{ role: 'user', content: [
            { type: 'image', source: { type: 'base64', media_type: flowImageData.mediaType, data: flowImageData.base64 } },
            { type: 'text', text: 'Extract all email flow/campaign data from this screenshot into the JSON format specified.' }
          ] }]
        })
      });
      const data = await resp.json();
      const text = data.content?.[0]?.text || '';
      const jsonMatch = text.match(/\[[\s\S]*\]/);
      if (!jsonMatch) throw new Error('Could not parse AI response');
      const rows = JSON.parse(jsonMatch[0]);
      setFlowAiPreview(rows);
    } catch (err) { setFlowError('AI parsing failed: ' + err.message); }
    finally { setFlowAiLoading(false); }
  };

  const handleFlowCsv = (file) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: (results) => {
        const missing = ['week', 'type', 'sends', 'revenue'].filter(h => !results.meta.fields?.includes(h));
        if (missing.length) { setFlowError(`Missing columns: ${missing.join(', ')}`); return; }
        setFlowCsvPreview(results.data);
      },
      error: (err) => setFlowError(err.message),
    });
  };

  const confirmImport = (rows) => {
    dispatch({ type: 'APPEND_DATA', source: 'emailFlows', payload: rows });
    setShowAddFlow(false); setFlowAiPreview(null); setFlowCsvPreview(null); setFlowImageData(null); setFlowImagePreview(null);
  };

  const { start, end } = state.dateRange;
  const compRange = useMemo(() => computeComparisonRange(start, end, state.comparison), [start, end, state.comparison]);
  const compLabel = COMP_LABELS[state.comparison] || null;
  const filtered = useMemo(() => filterByDateRange(state.emailFlows, start, end, 'week'), [state.emailFlows, start, end]);
  const period = state.tabPeriods.email || 'weekly';
  const data = period === 'monthly' ? aggregateEmailFlowsByMonth(filtered) : filtered;
  const weeks = [...new Set(data.map(r => r.week))].sort();
  const latestWeek = weeks[weeks.length - 1];
  const latestData = data.filter(r => r.week === latestWeek);
  const latestCampaign = latestData.find(r => r.type === 'Campaign');
  const totalEmailRev = latestData.reduce((s, r) => s + (r.revenue || 0), 0);
  const flowRev = latestData.filter(r => r.type === 'Flow').reduce((s, r) => s + (r.revenue || 0), 0);
  const fRevenue = useMemo(() => filterByDateRange(state.revenue, start, end, 'week'), [state.revenue, start, end]);
  const totalRev = fRevenue.length > 0 ? fRevenue[fRevenue.length - 1].totalRevenue : 1;
  const emailPct = (totalEmailRev / totalRev) * 100;

  // Channel cost & ROI from state.channelCosts
  const totalEmailRevRange = filtered.reduce((s, r) => s + (r.revenue || 0), 0);
  const emailCostROI = getChannelCostAndROI(state.channelCosts, 'Emails', totalEmailRevRange, start, end);

  // Comparison period
  const compEmailFlows = useMemo(() => compRange ? filterByDateRange(state.emailFlows, compRange.start, compRange.end, 'week') : [], [state.emailFlows, compRange]);
  const compEmail = useMemo(() => {
    if (!compRange || !compEmailFlows.length) return null;
    const cData = period === 'monthly' ? aggregateEmailFlowsByMonth(compEmailFlows) : compEmailFlows;
    const cWeeks = [...new Set(cData.map(r => r.week))].sort();
    const cLatest = cWeeks[cWeeks.length - 1];
    const cLatestData = cData.filter(r => r.week === cLatest);
    const cCampaign = cLatestData.find(r => r.type === 'Campaign');
    const cTotalRev = cLatestData.reduce((s, r) => s + (r.revenue || 0), 0);
    const cFlowRev = cLatestData.filter(r => r.type === 'Flow').reduce((s, r) => s + (r.revenue || 0), 0);
    const cRevData = filterByDateRange(state.revenue, compRange.start, compRange.end, 'week');
    const cBizRev = cRevData.length > 0 ? cRevData[cRevData.length - 1].totalRevenue : 1;
    return { totalRev: cTotalRev, flowRev: cFlowRev, emailPct: (cTotalRev / cBizRev) * 100, listSize: cCampaign?.listSize, openRate: cCampaign?.openRate, unsubRate: cCampaign?.unsubRate };
  }, [compRange, compEmailFlows, period, state.revenue]);

  // Build comparison weekly totals for overlay line
  const compWeeklyTotals = useMemo(() => {
    if (!compRange || !compEmailFlows.length) return [];
    const cData = period === 'monthly' ? aggregateEmailFlowsByMonth(compEmailFlows) : compEmailFlows;
    const cWeeks = [...new Set(cData.map(r => r.week))].sort();
    return cWeeks.map(w => cData.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0));
  }, [compRange, compEmailFlows, period]);

  const weeklyData = weeks.map((w, i) => {
    const rows = data.filter(r => r.week === w);
    const camp = rows.find(r => r.type === 'Campaign');
    const flowR = rows.filter(r => r.type === 'Flow').reduce((s, r) => s + (r.revenue || 0), 0);
    const entry = { week: w.slice(5), campaignRevenue: camp?.revenue || 0, flowRevenue: flowR };
    if (compWeeklyTotals.length > 0) entry.prevTotalRev = compWeeklyTotals[i] ?? null;
    return entry;
  });

  const flowNames = [...new Set(data.filter(r => r.type === 'Flow').map(r => r.flowName))];
  const flowTotalRev = flowNames.map(fn => ({
    name: fn,
    revenue: data.filter(r => r.flowName === fn).reduce((s, r) => s + (r.revenue || 0), 0),
    fill: CRM_CHANNEL_COLORS[fn] || C.info
  })).sort((a, b) => b.revenue - a.revenue);

  const campaignTrends = weeks.map(w => {
    const c = data.find(r => r.week === w && r.type === 'Campaign');
    return { week: w.slice(5), openRate: c?.openRate || 0, ctr: c?.ctr || 0 };
  });

  const last4Weeks = weeks.slice(-4);
  const flowPerformance = flowNames.map(fn => {
    const rows = data.filter(r => r.flowName === fn && last4Weeks.includes(r.week));
    const totalSends = rows.reduce((s, r) => s + r.sends, 0);
    const totalOpens = rows.reduce((s, r) => s + r.opens, 0);
    const totalClicks = rows.reduce((s, r) => s + r.clicks, 0);
    const totalUnsubs = rows.reduce((s, r) => s + r.unsubscribes, 0);
    const totalRev = rows.reduce((s, r) => s + r.revenue, 0);
    return {
      name: fn,
      sends: totalSends,
      openRate: totalSends > 0 ? (totalOpens / totalSends) * 100 : 0,
      ctr: totalSends > 0 ? (totalClicks / totalSends) * 100 : 0,
      unsubRate: totalSends > 0 ? (totalUnsubs / totalSends) * 100 : 0,
      revenue: totalRev,
      revPerSend: totalSends > 0 ? totalRev / totalSends : 0,
      isNew: fn === 'Re-Engagement 90d',
    };
  }).sort((a, b) => b.revenue - a.revenue);

  const waterfallData = weeks.map(w => {
    const row = { week: w.slice(5) };
    data.filter(r => r.week === w).forEach(r => {
      const key = r.type === 'Campaign' ? 'Campaign' : r.flowName;
      row[key] = r.revenue;
    });
    return row;
  });
  const waterfallKeys = ['Campaign', ...flowNames];

  const periodLabel = period === 'monthly' ? 'Month' : 'Week';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
        <TimePeriodToggle tab="email" tabPeriods={state.tabPeriods} dispatch={dispatch} />
        <button onClick={() => { setShowAddFlow(!showAddFlow); setFlowError(null); setFlowAiPreview(null); setFlowCsvPreview(null); }} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: showAddFlow ? C.textTertiary : C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>
          {showAddFlow ? 'Cancel' : '+ Add Flow'}
        </button>
      </div>

      {showAddFlow && (
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `2px solid ${C.primary}` }}>
          <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
            {['manual', 'screenshot', 'csv'].map(m => (
              <button key={m} onClick={() => { setAddFlowMode(m); setFlowError(null); }} style={{ padding: '6px 16px', borderRadius: 4, border: `1px solid ${addFlowMode === m ? C.primary : C.cardBorder}`, background: addFlowMode === m ? C.primary : 'transparent', color: addFlowMode === m ? '#fff' : C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer', textTransform: 'capitalize' }}>{m}</button>
            ))}
          </div>
          {flowError && <div style={{ background: '#FDE8E8', borderRadius: 4, padding: 8, fontSize: 12, color: '#D81F26', marginBottom: 12 }}>{flowError}</div>}

          {addFlowMode === 'manual' && (
            <div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                <input type="date" value={flowForm.week} onChange={e => setFlowForm({ ...flowForm, week: e.target.value })} style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <select value={flowForm.type} onChange={e => setFlowForm({ ...flowForm, type: e.target.value })} style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }}>
                  <option value="Flow">Flow</option><option value="Campaign">Campaign</option>
                </select>
                <input value={flowForm.flowName} onChange={e => setFlowForm({ ...flowForm, flowName: e.target.value })} placeholder="Flow name" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <input type="number" value={flowForm.sends} onChange={e => setFlowForm({ ...flowForm, sends: e.target.value })} placeholder="Sends" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <input type="number" step="0.1" value={flowForm.openRate} onChange={e => setFlowForm({ ...flowForm, openRate: e.target.value })} placeholder="Open Rate %" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <input type="number" step="0.1" value={flowForm.ctr} onChange={e => setFlowForm({ ...flowForm, ctr: e.target.value })} placeholder="CTR %" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <input type="number" step="0.01" value={flowForm.unsubRate} onChange={e => setFlowForm({ ...flowForm, unsubRate: e.target.value })} placeholder="Unsub Rate %" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
                <input type="number" value={flowForm.revenue} onChange={e => setFlowForm({ ...flowForm, revenue: e.target.value })} placeholder="Revenue (£)" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }} />
              </div>
              <button onClick={submitManualFlow} style={{ marginTop: 12, padding: '8px 24px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Add Flow</button>
            </div>
          )}

          {addFlowMode === 'screenshot' && (
            <div>
              <input ref={flowImageRef} type="file" accept="image/*" onChange={handleFlowImage} style={{ display: 'none' }} />
              {!flowImagePreview ? (
                <div onClick={() => flowImageRef.current?.click()} style={{ border: `2px dashed ${C.cardBorder}`, borderRadius: 6, padding: 30, textAlign: 'center', cursor: 'pointer', color: C.textTertiary, fontSize: 13 }}>
                  Click or drag to upload a screenshot
                </div>
              ) : (
                <div>
                  <img src={flowImagePreview} alt="Preview" style={{ maxHeight: 200, borderRadius: 4, marginBottom: 12 }} />
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button onClick={organizeFlowScreenshot} disabled={flowAiLoading} style={{ padding: '8px 24px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', opacity: flowAiLoading ? 0.6 : 1 }}>
                      {flowAiLoading ? 'Parsing...' : 'Extract with AI'}
                    </button>
                    <button onClick={() => { setFlowImageData(null); setFlowImagePreview(null); setFlowAiPreview(null); }} style={{ padding: '8px 16px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, cursor: 'pointer' }}>Clear</button>
                  </div>
                </div>
              )}
              {flowAiPreview && (
                <div style={{ marginTop: 12 }}>
                  <p style={{ fontSize: 12, color: C.textSecondary, marginBottom: 8 }}>Found {flowAiPreview.length} row(s):</p>
                  <div style={{ overflowX: 'auto', maxHeight: 200 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
                      <thead><tr>{Object.keys(flowAiPreview[0] || {}).map(h => <th key={h} style={{ padding: '4px 8px', textAlign: 'left', borderBottom: `1px solid ${C.cardBorder}`, color: C.textSecondary, fontSize: 10 }}>{h}</th>)}</tr></thead>
                      <tbody>{flowAiPreview.map((r, i) => <tr key={i}>{Object.values(r).map((v, j) => <td key={j} style={{ padding: '4px 8px', borderBottom: `1px solid ${C.divider}`, color: C.textPrimary }}>{String(v)}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                  <button onClick={() => confirmImport(flowAiPreview)} style={{ marginTop: 8, padding: '8px 24px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Confirm & Import</button>
                </div>
              )}
            </div>
          )}

          {addFlowMode === 'csv' && (
            <div>
              <input ref={flowCsvRef} type="file" accept=".csv" onChange={e => e.target.files?.[0] && handleFlowCsv(e.target.files[0])} style={{ display: 'none' }} />
              {!flowCsvPreview ? (
                <div onClick={() => flowCsvRef.current?.click()} style={{ border: `2px dashed ${C.cardBorder}`, borderRadius: 6, padding: 30, textAlign: 'center', cursor: 'pointer', color: C.textTertiary, fontSize: 13 }}>
                  Click to upload a CSV file<br /><span style={{ fontSize: 11 }}>Required: week, type, sends, revenue</span>
                </div>
              ) : (
                <div>
                  <p style={{ fontSize: 12, color: C.textSecondary, marginBottom: 8 }}>Preview: {flowCsvPreview.length} row(s)</p>
                  <div style={{ overflowX: 'auto', maxHeight: 200 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
                      <thead><tr>{Object.keys(flowCsvPreview[0] || {}).map(h => <th key={h} style={{ padding: '4px 8px', textAlign: 'left', borderBottom: `1px solid ${C.cardBorder}`, color: C.textSecondary, fontSize: 10 }}>{h}</th>)}</tr></thead>
                      <tbody>{flowCsvPreview.slice(0, 5).map((r, i) => <tr key={i}>{Object.values(r).map((v, j) => <td key={j} style={{ padding: '4px 8px', borderBottom: `1px solid ${C.divider}`, color: C.textPrimary }}>{String(v ?? '')}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                  <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                    <button onClick={() => confirmImport(flowCsvPreview)} style={{ padding: '8px 24px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Confirm & Import</button>
                    <button onClick={() => { setFlowCsvPreview(null); }} style={{ padding: '8px 16px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, cursor: 'pointer' }}>Clear</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label={`Total Email Revenue (${periodLabel})`} value={totalEmailRev} format="currency" status="good" compDelta={compEmail ? pctChange(totalEmailRev, compEmail.totalRev) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Email % of Revenue" value={emailPct} format="percent" status={emailPct < 15 ? 'warning' : 'good'} compDelta={compEmail ? pctChange(emailPct, compEmail.emailPct) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="List Size" value={latestCampaign?.listSize} format="number" status="good" compDelta={compEmail ? pctChange(latestCampaign?.listSize, compEmail.listSize) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Avg Campaign Open Rate" value={latestCampaign?.openRate} format="percent" status={latestCampaign?.openRate < 35 ? 'danger' : latestCampaign?.openRate < 40 ? 'warning' : 'good'} compDelta={compEmail ? pctChange(latestCampaign?.openRate, compEmail.openRate) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label={`Flow Revenue (${periodLabel})`} value={flowRev} format="currency" status="good" compDelta={compEmail ? pctChange(flowRev, compEmail.flowRev) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Unsubscribe Rate" value={latestCampaign?.unsubRate} format="percent" status={latestCampaign?.unsubRate > 0.5 ? 'danger' : 'good'} compDelta={compEmail ? pctChange(latestCampaign?.unsubRate, compEmail.unsubRate) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Email Channel Cost" value={emailCostROI.totalCost} format="currency" status="neutral" presentationMode={presentationMode} />
        <KPICard label="Email ROI" value={emailCostROI.roi} format="multiplier" status={emailCostROI.roi > 3 ? 'good' : emailCostROI.roi > 1 ? 'warning' : 'danger'} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Flow vs Campaign Revenue" tooltip="Compares automated flow revenue against one-off campaign revenue per week." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={weeklyData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="campaignRevenue" stackId="a" fill="#124A2B" name="Campaign" radius={[0,0,0,0]} />
            <Bar dataKey="flowRevenue" stackId="a" fill={C.info} name="Flows" radius={[4,4,0,0]} />
            {compRange && <Line type="monotone" dataKey="prevTotalRev" stroke={C.warning} strokeDasharray="6 3" strokeWidth={2} dot={false} name={`Total Rev (${compLabel})`} connectNulls />}
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Revenue by Flow" tooltip="Total revenue generated by each automated flow within the selected date range." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={flowTotalRev} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
              <YAxis type="category" dataKey="name" width={140} tick={{ fontSize: 11, fill: C.textSecondary }} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
              <Bar dataKey="revenue" radius={[0,4,4,0]}>
                {flowTotalRev.map((e, i) => <Cell key={i} fill={e.fill} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Campaign Open Rate & CTR Trends" tooltip="Weekly trends for campaign open rates and click-through rates." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <LineChart data={campaignTrends}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
              <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
              <Tooltip content={<ChartTooltip formatter={(v) => `${v.toFixed(1)}%`} />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="openRate" stroke={C.primary} strokeWidth={2} dot={{ r: 3 }} name="Open Rate" />
              <Line type="monotone" dataKey="ctr" stroke={C.success} strokeWidth={2} dot={{ r: 3 }} name="CTR" />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Flow Performance (Last 4 Weeks)" tooltip="Per-flow breakdown of sends, opens, clicks, and revenue for the most recent 4 weeks." />
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Flow','Sends','Open Rate','CTR','Unsub Rate','Revenue','Rev/Send'].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Flow' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {flowPerformance.map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                  <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>
                    {r.name} {r.isNew && <span style={{ background: C.primary, color: '#fff', fontSize: 9, padding: '2px 6px', borderRadius: 4, marginLeft: 6, fontWeight: 700 }}>NEW</span>}
                  </td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.sends)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatPercent(r.openRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatPercent(r.ctr)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: r.unsubRate > 1.0 ? C.danger : C.textSecondary }}>{formatPercent(r.unsubRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(r.revenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrencyDecimal(r.revPerSend)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Weekly Email Revenue by Source" tooltip="Stacked bar chart showing weekly email revenue split by individual flows and campaigns." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <AreaChart data={waterfallData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {waterfallKeys.map(k => (
              <Area key={k} type="monotone" dataKey={k} stackId="1" fill={CRM_CHANNEL_COLORS[k] || '#124A2B'} stroke={CRM_CHANNEL_COLORS[k] || '#124A2B'} fillOpacity={0.7} />
            ))}
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

// ─── CHANNEL FLOW SECTION (reusable for WhatsApp & Postcard) ───
function ChannelFlowSection({ state, presentationMode, dispatch, channelKey, channelLabel, channelColor, dataKey }) {
  const chartHeight = presentationMode ? 400 : 300;
  const { start, end } = state.dateRange;
  const [showAddFlow, setShowAddFlow] = useState(false);
  const [flowForm, setFlowForm] = useState({ week: '', flowName: '', sends: '', delivered: '', responses: '', conversions: '', revenue: '', cost: '' });
  const [flowError, setFlowError] = useState(null);

  const data = useMemo(() => filterByDateRange(state[dataKey] || [], start, end, 'week'), [state[dataKey], start, end]);

  // KPIs
  const totalRevenue = data.reduce((s, r) => s + (r.revenue || 0), 0);
  const totalSends = data.reduce((s, r) => s + (r.sends || 0), 0);
  const totalResponses = data.reduce((s, r) => s + (r.responses || 0), 0);
  const totalConversions = data.reduce((s, r) => s + (r.conversions || 0), 0);
  const totalCost = data.reduce((s, r) => s + (r.cost || 0), 0);
  const avgResponseRate = totalSends > 0 ? (totalResponses / totalSends) * 100 : 0;
  const avgConvRate = totalSends > 0 ? (totalConversions / totalSends) * 100 : 0;

  // Platform-level cost from channelCosts
  const platformCostROI = getChannelCostAndROI(state.channelCosts, channelLabel, totalRevenue, start, end);

  // Weekly trend
  const weeks = [...new Set(data.map(r => r.week))].sort();
  const weeklyTrend = weeks.map(w => {
    const wd = data.filter(r => r.week === w);
    return { week: w.slice(5), revenue: wd.reduce((s, r) => s + (r.revenue || 0), 0), sends: wd.reduce((s, r) => s + (r.sends || 0), 0), conversions: wd.reduce((s, r) => s + (r.conversions || 0), 0) };
  });

  // Flow breakdown
  const flowNames = [...new Set(data.map(r => r.flowName))].sort();
  const flowSummary = flowNames.map(fn => {
    const fd = data.filter(r => r.flowName === fn);
    const s = fd.reduce((a, r) => a + (r.sends || 0), 0);
    const resp = fd.reduce((a, r) => a + (r.responses || 0), 0);
    const conv = fd.reduce((a, r) => a + (r.conversions || 0), 0);
    const rev = fd.reduce((a, r) => a + (r.revenue || 0), 0);
    const c = fd.reduce((a, r) => a + (r.cost || 0), 0);
    return { flowName: fn, sends: s, responses: resp, responseRate: s > 0 ? ((resp / s) * 100).toFixed(1) : '0', conversions: conv, convRate: s > 0 ? ((conv / s) * 100).toFixed(1) : '0', revenue: rev, cost: c, revPerSend: s > 0 ? (rev / s).toFixed(2) : '0' };
  }).sort((a, b) => b.revenue - a.revenue);

  // Revenue by flow chart
  const flowChartData = flowSummary.map(f => ({ name: f.flowName, revenue: f.revenue }));

  // Weekly by flow (stacked)
  const FLOW_COLORS = ['#124A2B', '#18917B', '#3B82F6', '#F59E0B', '#D81F26', '#7C3AED'];
  const weeklyByFlow = weeks.map(w => {
    const entry = { week: w.slice(5) };
    flowNames.forEach(fn => { entry[fn] = data.filter(r => r.week === w && r.flowName === fn).reduce((s, r) => s + (r.revenue || 0), 0); });
    return entry;
  });

  const handleAddFlow = () => {
    const { week, flowName, sends, delivered, responses, conversions, revenue, cost } = flowForm;
    if (!week || !flowName) { setFlowError('Week and flow name are required'); return; }
    const row = { week, flowName, sends: +sends || 0, delivered: +delivered || 0, responses: +responses || 0, conversions: +conversions || 0, revenue: +revenue || 0, cost: +cost || 0 };
    dispatch({ type: 'APPEND_DATA', source: dataKey, payload: [row] });
    setFlowForm({ week: '', flowName: '', sends: '', delivered: '', responses: '', conversions: '', revenue: '', cost: '' });
    setFlowError(null);
    setShowAddFlow(false);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label={`${channelLabel} Revenue`} value={totalRevenue} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="Total Sends" value={totalSends} format="number" status="good" presentationMode={presentationMode} />
        <KPICard label="Response Rate" value={avgResponseRate} format="percent" status={avgResponseRate > 20 ? 'good' : 'warning'} presentationMode={presentationMode} />
        <KPICard label="Conversion Rate" value={avgConvRate} format="percent" status={avgConvRate > 2 ? 'good' : 'warning'} presentationMode={presentationMode} />
        <KPICard label="Channel Cost" value={platformCostROI.totalCost} format="currency" status="neutral" presentationMode={presentationMode} />
        <KPICard label="Channel ROI" value={platformCostROI.roi} format="multiplier" status={platformCostROI.roi > 3 ? 'good' : platformCostROI.roi > 1 ? 'warning' : 'danger'} presentationMode={presentationMode} />
      </div>

      {/* Revenue Over Time */}
      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title={`${channelLabel} Revenue Over Time`} tooltip={`Weekly revenue from ${channelLabel} flows.`} />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={weeklyTrend}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(1)}k`} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <Tooltip content={<ChartTooltip formatter={(v, name) => name === 'Conversions' ? formatNumber(v) : formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="revenue" yAxisId="left" fill={channelColor} name="Revenue" radius={[4, 4, 0, 0]} />
            <Line type="monotone" dataKey="conversions" yAxisId="right" stroke="#F59E0B" strokeWidth={2} dot={false} name="Conversions" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
        {/* Revenue by Flow */}
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Revenue by Flow" tooltip="Total revenue per flow." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={flowChartData} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(1)}k`} />
              <YAxis dataKey="name" type="category" width={130} tick={{ fontSize: 11, fill: C.textTertiary }} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
              <Bar dataKey="revenue" fill={channelColor} name="Revenue" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Weekly Revenue by Flow (stacked) */}
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Weekly Revenue by Flow" tooltip="Stacked weekly revenue breakdown by flow." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <AreaChart data={weeklyByFlow}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
              <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(1)}k`} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              {flowNames.map((fn, i) => (
                <Area key={fn} type="monotone" dataKey={fn} stackId="flows" fill={FLOW_COLORS[i % FLOW_COLORS.length]} stroke={FLOW_COLORS[i % FLOW_COLORS.length]} fillOpacity={0.6} name={fn} />
              ))}
            </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Flow Performance Table */}
      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <ChartHeader title="Flow Performance" tooltip="Detailed performance metrics per flow." />
          <button onClick={() => setShowAddFlow(!showAddFlow)} style={{ background: channelColor, color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', cursor: 'pointer', fontWeight: 600, fontSize: 13 }}>
            {showAddFlow ? 'Cancel' : '+ Add Flow'}
          </button>
        </div>

        {showAddFlow && (
          <div style={{ background: C.bg, borderRadius: 6, padding: 16, marginBottom: 16, border: `1px solid ${C.cardBorder}` }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(160px, 1fr))', gap: 10, marginBottom: 12 }}>
              {[
                { key: 'week', label: 'Week (YYYY-MM-DD)', type: 'date' },
                { key: 'flowName', label: 'Flow Name', type: 'text' },
                { key: 'sends', label: 'Sends', type: 'number' },
                { key: 'delivered', label: 'Delivered', type: 'number' },
                { key: 'responses', label: 'Responses', type: 'number' },
                { key: 'conversions', label: 'Conversions', type: 'number' },
                { key: 'revenue', label: 'Revenue (£)', type: 'number' },
                { key: 'cost', label: 'Cost (£)', type: 'number' },
              ].map(f => (
                <div key={f.key}>
                  <label style={{ fontSize: 11, color: C.textSecondary, display: 'block', marginBottom: 4 }}>{f.label}</label>
                  <input type={f.type} value={flowForm[f.key]} onChange={e => setFlowForm(p => ({ ...p, [f.key]: e.target.value }))}
                    style={{ width: '100%', padding: '6px 10px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary, boxSizing: 'border-box' }} />
                </div>
              ))}
            </div>
            {flowError && <div style={{ color: '#D81F26', fontSize: 12, marginBottom: 8 }}>{flowError}</div>}
            <button onClick={handleAddFlow} style={{ background: channelColor, color: '#fff', border: 'none', borderRadius: 6, padding: '8px 20px', cursor: 'pointer', fontWeight: 600, fontSize: 13 }}>Save Flow</button>
          </div>
        )}

        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Flow','Sends','Responses','Resp. Rate','Conversions','Conv. Rate','Revenue','Cost','Rev/Send'].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Flow' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {flowSummary.map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                  <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{r.flowName}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.sends)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.responses)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{r.responseRate}%</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.conversions)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{r.convRate}%</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(r.revenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.cost)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>£{r.revPerSend}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function WhatsAppSection({ state, presentationMode, dispatch }) {
  return <ChannelFlowSection state={state} presentationMode={presentationMode} dispatch={dispatch} channelKey="whatsapp" channelLabel="WhatsApp" channelColor="#25D366" dataKey="whatsappFlows" />;
}

function PostcardSection({ state, presentationMode, dispatch }) {
  return <ChannelFlowSection state={state} presentationMode={presentationMode} dispatch={dispatch} channelKey="postcard" channelLabel="Postcards" channelColor="#F59E0B" dataKey="postcardFlows" />;
}

function LoyaltySection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const { start, end } = state.dateRange;
  const compRange = useMemo(() => computeComparisonRange(start, end, state.comparison), [start, end, state.comparison]);
  const compLabel = COMP_LABELS[state.comparison] || null;
  const data = useMemo(() => filterByDateRange(state.loyalty, start, end, 'month'), [state.loyalty, start, end]);
  const latest = data.length > 0 ? data[data.length - 1] : null;
  const aovLift = latest ? ((latest.memberAOV - latest.nonMemberAOV) / latest.nonMemberAOV) * 100 : 0;
  const tier6thLTV = latest?.tier6thOrderLTV || 0;
  const nonMemLTV = latest?.nonMemberLTV || 0;
  const ltvLift = nonMemLTV > 0 ? ((tier6thLTV - nonMemLTV) / nonMemLTV) * 100 : 0;

  // Comparison period
  const compData = useMemo(() => compRange ? filterByDateRange(state.loyalty, compRange.start, compRange.end, 'month') : [], [state.loyalty, compRange]);
  const compLoy = useMemo(() => {
    if (!compRange || !compData.length) return null;
    const cl = compData[compData.length - 1];
    const cAovLift = cl.nonMemberAOV ? ((cl.memberAOV - cl.nonMemberAOV) / cl.nonMemberAOV) * 100 : 0;
    const cLtvLift = cl.nonMemberLTV > 0 ? ((cl.tier6thOrderLTV - cl.nonMemberLTV) / cl.nonMemberLTV) * 100 : 0;
    return { totalMembers: cl.totalMembers, newEnrollments: cl.newEnrollments, redemptionRate: cl.redemptionRate, memberAOV: cl.memberAOV, nonMemberAOV: cl.nonMemberAOV, aovLift: cAovLift, tier6thOrderLTV: cl.tier6thOrderLTV, nonMemberLTV: cl.nonMemberLTV, ltvLift: cLtvLift };
  }, [compRange, compData]);

  // Per-product milestone data
  const [selectedProduct, setSelectedProduct] = useState(null);
  const [showAddProduct, setShowAddProduct] = useState(false);
  const [addForm, setAddForm] = useState({ product: '', month: '', tier2nd: '', tier2ndAOV: '', tier2ndLTV: '', tier3rd: '', tier3rdAOV: '', tier3rdLTV: '', tier6th: '', tier6thAOV: '', tier6thLTV: '' });

  const productData = useMemo(() => filterByDateRange(state.milestoneProducts || [], start, end, 'month'), [state.milestoneProducts, start, end]);
  const productNames = useMemo(() => [...new Set(productData.map(r => r.product))].sort(), [productData]);

  const chartData = data.map((m, i) => {
    const entry = {
      month: m.month.slice(2),
      totalMembers: m.totalMembers,
      newEnrollments: m.newEnrollments,
      pointsIssued: m.pointsIssued,
      pointsRedeemed: m.pointsRedeemed,
      redemptionRate: m.redemptionRate,
      revenueFromMembers: m.revenueFromMembers,
      revenueFromNonMembers: m.revenueFromNonMembers,
      memberAOV: m.memberAOV,
      nonMemberAOV: m.nonMemberAOV,
      rewardCostGBP: m.rewardCostGBP,
      memberRetentionRate: m.memberRetentionRate,
      nonMemberRetentionRate: m.nonMemberRetentionRate,
      tier2ndOrderAOV: m.tier2ndOrderAOV,
      tier3rdOrderAOV: m.tier3rdOrderAOV,
      tier6thOrderAOV: m.tier6thOrderAOV,
      tier2ndOrderLTV: m.tier2ndOrderLTV,
      tier3rdOrderLTV: m.tier3rdOrderLTV,
      tier6thOrderLTV: m.tier6thOrderLTV,
      nonMemberLTV: m.nonMemberLTV,
    };
    if (compData.length > 0 && compData[i]) entry.prevTotalMembers = compData[i].totalMembers;
    return entry;
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <TimePeriodToggle tab="loyalty" tabPeriods={state.tabPeriods} dispatch={dispatch} options={['monthly']} />
      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label="Total Members" value={latest?.totalMembers} format="number" status="good" sparkData={data} sparkKey="totalMembers" delta={!compLoy ? calcDelta(data, 'totalMembers') : null} compDelta={compLoy ? pctChange(latest?.totalMembers, compLoy.totalMembers) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="New Enrollments (Month)" value={latest?.newEnrollments} format="number" status="good" compDelta={compLoy ? pctChange(latest?.newEnrollments, compLoy.newEnrollments) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Redemption Rate" value={latest?.redemptionRate} format="percent" status={latest?.redemptionRate < 10 ? 'warning' : 'good'} sparkData={data} sparkKey="redemptionRate" compDelta={compLoy ? pctChange(latest?.redemptionRate, compLoy.redemptionRate) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Member AOV" value={latest?.memberAOV} format="currencyDecimal" status="good" compDelta={compLoy ? pctChange(latest?.memberAOV, compLoy.memberAOV) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Non-Member AOV" value={latest?.nonMemberAOV} format="currencyDecimal" status="neutral" compDelta={compLoy ? pctChange(latest?.nonMemberAOV, compLoy.nonMemberAOV) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="AOV Lift (Member vs Non)" value={aovLift} format="percent" status="good" compDelta={compLoy ? pctChange(aovLift, compLoy.aovLift) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="6th Order LTV" value={tier6thLTV} format="currencyDecimal" status="good" sparkData={data} sparkKey="tier6thOrderLTV" compDelta={compLoy ? pctChange(tier6thLTV, compLoy.tier6thOrderLTV) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="LTV Lift (6th vs Non)" value={ltvLift} format="percent" status={ltvLift >= 200 ? 'good' : 'warning'} compDelta={compLoy ? pctChange(ltvLift, compLoy.ltvLift) : null} compLabel={compLabel} presentationMode={presentationMode} />
      </div>

      {/* Milestone Tier Breakdown Cards */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 12 }}>
        {MILESTONE_TIERS.map(tier => {
          const members = latest?.[`tier${tier.key}Members`] || 0;
          const aov = latest?.[`tier${tier.key}AOV`] || 0;
          const ltv = latest?.[`tier${tier.key}LTV`] || 0;
          return (
            <div key={tier.key} style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.cardBorder}`, borderLeft: `4px solid ${TIER_COLORS[tier.label]}` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: TIER_COLORS[tier.label], marginBottom: 4 }}>{tier.label}</div>
              <div style={{ height: 10 }} />
              <div style={{ fontSize: 22, fontWeight: 700, color: C.textPrimary, marginBottom: 6 }}>{formatNumber(members)}</div>
              <div style={{ fontSize: 11, color: C.textSecondary, marginBottom: 2 }}>members</div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 8 }}>
                <div>
                  <div style={{ fontSize: 10, color: C.textTertiary }}>AOV</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>{formatCurrencyDecimal(aov)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, color: C.textTertiary }}>LTV</div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.success }}>{formatCurrencyDecimal(ltv)}</div>
                </div>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Milestone Membership Growth" tooltip="Monthly total milestone reward members (bars) and new enrolments (line) over time." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <Tooltip content={<ChartTooltip formatter={(v, name) => name.includes('Rate') ? formatPercent(v) : formatNumber(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Area yAxisId="left" type="monotone" dataKey="totalMembers" fill={C.primary} stroke={C.primary} fillOpacity={0.2} name="Total Members" />
            <Bar yAxisId="right" dataKey="newEnrollments" fill={C.success} name="New Enrollments" radius={[4,4,0,0]} />
            {compRange && <Line yAxisId="left" type="monotone" dataKey="prevTotalMembers" stroke={C.warning} strokeDasharray="6 3" strokeWidth={2} dot={false} name={`Members (${compLabel})`} connectNulls />}
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Points Economy" tooltip="Monthly points issued vs redeemed, showing the balance of the milestone reward currency." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <ComposedChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
              <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${(v/1000).toFixed(0)}k`} />
              <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
              <Tooltip content={<ChartTooltip formatter={(v, name) => name === 'Redemption Rate' ? formatPercent(v) : formatNumber(v)} />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar yAxisId="left" dataKey="pointsIssued" fill={C.info} name="Points Issued" radius={[4,4,0,0]} />
              <Bar yAxisId="left" dataKey="pointsRedeemed" fill={C.success} name="Points Redeemed" radius={[4,4,0,0]} />
              <Line yAxisId="right" type="monotone" dataKey="redemptionRate" stroke={C.warning} strokeWidth={2} dot={{ r: 3 }} name="Redemption Rate" />
            </ComposedChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="AOV by Milestone Tier" tooltip="Average order value comparison across milestone tiers and non-members over time." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
              <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v}`} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrencyDecimal(v)} />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="tier2ndOrderAOV" fill={TIER_COLORS['2nd Order']} name="2nd Order" radius={[4,4,0,0]} />
              <Bar dataKey="tier3rdOrderAOV" fill={TIER_COLORS['3rd Order']} name="3rd Order" radius={[4,4,0,0]} />
              <Bar dataKey="tier6thOrderAOV" fill={TIER_COLORS['6th Order']} name="6th Order" radius={[4,4,0,0]} />
              <Bar dataKey="nonMemberAOV" fill="#94A3B8" name="Non-Member" radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="LTV by Milestone Tier" tooltip="Lifetime value comparison across milestone tiers and non-members, showing the core business case for the milestone program." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v}`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrencyDecimal(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Line type="monotone" dataKey="tier2ndOrderLTV" stroke={TIER_COLORS['2nd Order']} strokeWidth={2} dot={{ r: 4 }} name="2nd Order LTV" />
            <Line type="monotone" dataKey="tier3rdOrderLTV" stroke={TIER_COLORS['3rd Order']} strokeWidth={2} dot={{ r: 4 }} name="3rd Order LTV" />
            <Line type="monotone" dataKey="tier6thOrderLTV" stroke={TIER_COLORS['6th Order']} strokeWidth={2} dot={{ r: 4 }} name="6th Order LTV" />
            <Line type="monotone" dataKey="nonMemberLTV" stroke="#94A3B8" strokeWidth={2} dot={{ r: 4 }} name="Non-Member LTV" />
          </LineChart>
        </ResponsiveContainer>
      </div>

      {/* ─── Product Breakdown ─── */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
        <h3 style={{ fontSize: 16, fontWeight: 700, color: C.textPrimary, margin: 0 }}>Product Breakdown</h3>
        <div style={{ display: 'flex', gap: 8 }}>
          {selectedProduct && (
            <button onClick={() => setSelectedProduct(null)} style={{ padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>← All Products</button>
          )}
          <button onClick={() => { setShowAddProduct(!showAddProduct); setAddForm({ product: '', month: data.length ? data[data.length - 1].month : '2025-03', tier2nd: '', tier2ndAOV: '', tier2ndLTV: '', tier3rd: '', tier3rdAOV: '', tier3rdLTV: '', tier6th: '', tier6thAOV: '', tier6thLTV: '' }); }} style={{ padding: '6px 14px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            {showAddProduct ? 'Cancel' : '+ Add Product'}
          </button>
        </div>
      </div>

      {/* Add Product Form */}
      {showAddProduct && (
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.primary}40` }}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(140px, 1fr))', gap: 10 }}>
            <div>
              <label style={{ fontSize: 10, fontWeight: 600, color: C.textTertiary, textTransform: 'uppercase' }}>Product Name</label>
              <input value={addForm.product} onChange={e => setAddForm(f => ({ ...f, product: e.target.value }))} placeholder="e.g. Omega Oil" style={{ width: '100%', padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.surface, color: C.textPrimary, boxSizing: 'border-box' }} />
            </div>
            <div>
              <label style={{ fontSize: 10, fontWeight: 600, color: C.textTertiary, textTransform: 'uppercase' }}>Month</label>
              <input value={addForm.month} onChange={e => setAddForm(f => ({ ...f, month: e.target.value }))} placeholder="2025-03" style={{ width: '100%', padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.surface, color: C.textPrimary, boxSizing: 'border-box' }} />
            </div>
            {MILESTONE_TIERS.map(tier => (
              <React.Fragment key={tier.key}>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600, color: TIER_COLORS[tier.label] }}>{tier.label} Members</label>
                  <input type="number" value={addForm[`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}`]} onChange={e => setAddForm(f => ({ ...f, [`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}`]: e.target.value }))} style={{ width: '100%', padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.surface, color: C.textPrimary, boxSizing: 'border-box' }} />
                </div>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600, color: TIER_COLORS[tier.label] }}>{tier.label} AOV</label>
                  <input type="number" step="0.01" value={addForm[`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}AOV`]} onChange={e => setAddForm(f => ({ ...f, [`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}AOV`]: e.target.value }))} style={{ width: '100%', padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.surface, color: C.textPrimary, boxSizing: 'border-box' }} />
                </div>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600, color: TIER_COLORS[tier.label] }}>{tier.label} LTV</label>
                  <input type="number" step="0.01" value={addForm[`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}LTV`]} onChange={e => setAddForm(f => ({ ...f, [`tier${tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th'}LTV`]: e.target.value }))} style={{ width: '100%', padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.surface, color: C.textPrimary, boxSizing: 'border-box' }} />
                </div>
              </React.Fragment>
            ))}
          </div>
          <button onClick={() => {
            if (!addForm.product.trim() || !addForm.month.trim()) return;
            const record = {
              month: addForm.month, product: addForm.product.trim(),
              tier2nd: Number(addForm.tier2nd) || 0, tier2ndAOV: Number(addForm.tier2ndAOV) || 0, tier2ndLTV: Number(addForm.tier2ndLTV) || 0,
              tier3rd: Number(addForm.tier3rd) || 0, tier3rdAOV: Number(addForm.tier3rdAOV) || 0, tier3rdLTV: Number(addForm.tier3rdLTV) || 0,
              tier6th: Number(addForm.tier6th) || 0, tier6thAOV: Number(addForm.tier6thAOV) || 0, tier6thLTV: Number(addForm.tier6thLTV) || 0,
            };
            dispatch({ type: 'APPEND_DATA', source: 'milestoneProducts', payload: [record] });
            setShowAddProduct(false);
          }} style={{ marginTop: 10, padding: '6px 18px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Save Product</button>
        </div>
      )}

      {/* Product Grid (summary) */}
      {!selectedProduct && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))', gap: 12 }}>
          {productNames.map(name => {
            const latest = productData.filter(r => r.product === name).sort((a, b) => a.month.localeCompare(b.month)).slice(-1)[0];
            if (!latest) return null;
            return (
              <div key={name} onClick={() => setSelectedProduct(name)} style={{ background: C.cardBg, borderRadius: 6, padding: 16, border: `1px solid ${C.cardBorder}`, cursor: 'pointer', transition: 'border-color 0.15s' }} onMouseEnter={e => e.currentTarget.style.borderColor = C.primary} onMouseLeave={e => e.currentTarget.style.borderColor = C.cardBorder}>
                <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary, marginBottom: 10 }}>{name}</div>
                {MILESTONE_TIERS.map(tier => {
                  const key = tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th';
                  return (
                    <div key={tier.key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '3px 0', borderLeft: `3px solid ${TIER_COLORS[tier.label]}`, paddingLeft: 8, marginBottom: 4 }}>
                      <span style={{ fontSize: 11, color: C.textSecondary }}>{tier.label}</span>
                      <span style={{ fontSize: 11, fontWeight: 600, color: C.textPrimary }}>{formatNumber(latest[`tier${key}`])} members</span>
                    </div>
                  );
                })}
                <div style={{ marginTop: 8, padding: '6px 8px', background: `${C.success}10`, borderRadius: 4, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: 10, color: C.textTertiary }}>6th Order LTV</span>
                  <span style={{ fontSize: 14, fontWeight: 700, color: C.success }}>{formatCurrencyDecimal(latest.tier6thLTV)}</span>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Expanded Product Detail */}
      {selectedProduct && (() => {
        const pRows = productData.filter(r => r.product === selectedProduct).sort((a, b) => a.month.localeCompare(b.month));
        const pLatest = pRows.length > 0 ? pRows[pRows.length - 1] : null;
        const pChartData = pRows.map(r => ({
          month: r.month.slice(2),
          tier2ndAOV: r.tier2ndAOV, tier3rdAOV: r.tier3rdAOV, tier6thAOV: r.tier6thAOV,
          tier2ndLTV: r.tier2ndLTV, tier3rdLTV: r.tier3rdLTV, tier6thLTV: r.tier6thLTV,
          tier2nd: r.tier2nd, tier3rd: r.tier3rd, tier6th: r.tier6th,
        }));
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <h4 style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary, margin: 0 }}>{selectedProduct}</h4>

            {/* Tier KPI mini-cards */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 12 }}>
              {MILESTONE_TIERS.map(tier => {
                const key = tier.key === '2ndOrder' ? '2nd' : tier.key === '3rdOrder' ? '3rd' : '6th';
                return (
                  <div key={tier.key} style={{ background: C.cardBg, borderRadius: 6, padding: 14, border: `1px solid ${C.cardBorder}`, borderLeft: `4px solid ${TIER_COLORS[tier.label]}` }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: TIER_COLORS[tier.label], marginBottom: 2 }}>{tier.label}</div>
                    <div style={{ fontSize: 10, color: C.textTertiary, marginBottom: 8 }}>{tier.product}</div>
                    <div style={{ fontSize: 20, fontWeight: 700, color: C.textPrimary }}>{formatNumber(pLatest?.[`tier${key}`] || 0)}</div>
                    <div style={{ fontSize: 10, color: C.textSecondary, marginBottom: 6 }}>members</div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <div><div style={{ fontSize: 10, color: C.textTertiary }}>AOV</div><div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>{formatCurrencyDecimal(pLatest?.[`tier${key}AOV`] || 0)}</div></div>
                      <div><div style={{ fontSize: 10, color: C.textTertiary }}>LTV</div><div style={{ fontSize: 13, fontWeight: 600, color: C.success }}>{formatCurrencyDecimal(pLatest?.[`tier${key}LTV`] || 0)}</div></div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Per-product charts */}
            <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
              <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
                <ChartHeader title={`${selectedProduct} — AOV by Tier`} tooltip={`Average order value per milestone tier for ${selectedProduct} over time.`} />
                <ResponsiveContainer width="100%" height={chartHeight}>
                  <LineChart data={pChartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
                    <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
                    <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v}`} />
                    <Tooltip content={<ChartTooltip formatter={(v) => formatCurrencyDecimal(v)} />} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Line type="monotone" dataKey="tier2ndAOV" stroke={TIER_COLORS['2nd Order']} strokeWidth={2} dot={{ r: 4 }} name="2nd Order" />
                    <Line type="monotone" dataKey="tier3rdAOV" stroke={TIER_COLORS['3rd Order']} strokeWidth={2} dot={{ r: 4 }} name="3rd Order" />
                    <Line type="monotone" dataKey="tier6thAOV" stroke={TIER_COLORS['6th Order']} strokeWidth={2} dot={{ r: 4 }} name="6th Order" />
                  </LineChart>
                </ResponsiveContainer>
              </div>

              <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
                <ChartHeader title={`${selectedProduct} — LTV by Tier`} tooltip={`Lifetime value per milestone tier for ${selectedProduct} over time.`} />
                <ResponsiveContainer width="100%" height={chartHeight}>
                  <LineChart data={pChartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
                    <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
                    <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v}`} />
                    <Tooltip content={<ChartTooltip formatter={(v) => formatCurrencyDecimal(v)} />} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Line type="monotone" dataKey="tier2ndLTV" stroke={TIER_COLORS['2nd Order']} strokeWidth={2} dot={{ r: 4 }} name="2nd Order" />
                    <Line type="monotone" dataKey="tier3rdLTV" stroke={TIER_COLORS['3rd Order']} strokeWidth={2} dot={{ r: 4 }} name="3rd Order" />
                    <Line type="monotone" dataKey="tier6thLTV" stroke={TIER_COLORS['6th Order']} strokeWidth={2} dot={{ r: 4 }} name="6th Order" />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        );
      })()}

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Member vs Non-Member Revenue" tooltip="Stacked revenue from milestone members vs non-members, with AOV lines for each group." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v}`} />
            <Tooltip content={<ChartTooltip formatter={(v, name) => name.includes('AOV') ? formatCurrencyDecimal(v) : formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="revenueFromMembers" stackId="rev" fill={C.primary} name="Member Revenue" />
            <Bar yAxisId="left" dataKey="revenueFromNonMembers" stackId="rev" fill={C.textTertiary} name="Non-Member Revenue" radius={[4,4,0,0]} />
            <Line yAxisId="right" type="monotone" dataKey="memberAOV" stroke={C.success} strokeWidth={2} dot={{ r: 3 }} name="Member AOV" />
            <Line yAxisId="right" type="monotone" dataKey="nonMemberAOV" stroke={C.warning} strokeWidth={2} dot={{ r: 3 }} name="Non-Member AOV" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Retention: Members vs Non-Members" tooltip="Monthly retention rate comparison between milestone reward members and non-members." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} domain={[60, 100]} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatPercent(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Line type="monotone" dataKey="memberRetentionRate" stroke={C.primary} strokeWidth={2} dot={{ r: 4 }} name="Member Retention" />
            <Line type="monotone" dataKey="nonMemberRetentionRate" stroke={C.danger} strokeWidth={2} dot={{ r: 4 }} name="Non-Member Retention" />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function SegmentsSection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const [showSegLinkForm, setShowSegLinkForm] = useState(false);
  const [editSegLink, setEditSegLink] = useState(null);
  const [segLinkForm, setSegLinkForm] = useState({ name: '', type: 'segment', klaviyo_url: '', description: '', member_count: '' });
  const connStr = getNeonConnection();
  const segUsername = state?.currentUser?.displayName || localStorage.getItem('crm_display_name') || 'Anonymous';

  useEffect(() => {
    const load = async () => {
      try {
        if (connStr) {
          const rows = await neonQuery(connStr, 'SELECT * FROM segment_links ORDER BY created_at DESC');
          dispatch({ type: 'SET_SEGMENT_LINKS', payload: rows });
        } else {
          const local = JSON.parse(localStorage.getItem('crm_segment_links') || '[]');
          if (local.length > 0) dispatch({ type: 'SET_SEGMENT_LINKS', payload: local });
        }
      } catch (_) {
        const local = JSON.parse(localStorage.getItem('crm_segment_links') || '[]');
        if (local.length > 0) dispatch({ type: 'SET_SEGMENT_LINKS', payload: local });
      }
    };
    load();
  }, [connStr, dispatch]);

  const saveSegmentLink = async () => {
    if (!segLinkForm.name.trim()) return;
    const entry = { ...segLinkForm, member_count: segLinkForm.member_count ? Number(segLinkForm.member_count) : 0 };
    try {
      if (connStr && editSegLink) {
        await neonQuery(connStr, 'UPDATE segment_links SET name=$1, type=$2, klaviyo_url=$3, description=$4, member_count=$5, updated_at=NOW() WHERE id=$6', [entry.name, entry.type, entry.klaviyo_url, entry.description, entry.member_count, editSegLink.id]);
        dispatch({ type: 'UPDATE_SEGMENT_LINK', payload: { id: editSegLink.id, ...entry } });
      } else if (connStr) {
        const rows = await neonQuery(connStr, 'INSERT INTO segment_links (name, type, klaviyo_url, description, member_count, created_by) VALUES ($1,$2,$3,$4,$5,$6) RETURNING *', [entry.name, entry.type, entry.klaviyo_url, entry.description, entry.member_count, segUsername]);
        dispatch({ type: 'ADD_SEGMENT_LINK', payload: rows[0] });
      } else {
        const newLink = { id: Date.now(), ...entry, created_by: segUsername, created_at: new Date().toISOString() };
        dispatch({ type: editSegLink ? 'UPDATE_SEGMENT_LINK' : 'ADD_SEGMENT_LINK', payload: editSegLink ? { id: editSegLink.id, ...entry } : newLink });
      }
    } catch (_) {
      const newLink = { id: Date.now(), ...entry, created_by: segUsername, created_at: new Date().toISOString() };
      dispatch({ type: editSegLink ? 'UPDATE_SEGMENT_LINK' : 'ADD_SEGMENT_LINK', payload: editSegLink ? { id: editSegLink.id, ...entry } : newLink });
    }
    const updatedLinks = editSegLink
      ? state.segmentLinks.map(s => s.id === editSegLink.id ? { ...s, ...entry } : s)
      : [...state.segmentLinks, { id: Date.now(), ...entry }];
    localStorage.setItem('crm_segment_links', JSON.stringify(updatedLinks));
    setShowSegLinkForm(false); setEditSegLink(null);
    setSegLinkForm({ name: '', type: 'segment', klaviyo_url: '', description: '', member_count: '' });
  };

  const deleteSegmentLink = async (id) => {
    try { if (connStr) await neonQuery(connStr, 'DELETE FROM segment_links WHERE id=$1', [id]); } catch (_) {}
    dispatch({ type: 'DELETE_SEGMENT_LINK', payload: id });
    localStorage.setItem('crm_segment_links', JSON.stringify(state.segmentLinks.filter(s => s.id !== id)));
  };

  const { start, end } = state.dateRange;
  const compRange = useMemo(() => computeComparisonRange(start, end, state.comparison), [start, end, state.comparison]);
  const compLabel = COMP_LABELS[state.comparison] || null;
  const data = useMemo(() => filterByDateRange(state.segments, start, end, 'month'), [state.segments, start, end]);
  const latest = data.length > 0 ? data[data.length - 1] : null;

  // Comparison period
  const compData = useMemo(() => compRange ? filterByDateRange(state.segments, compRange.start, compRange.end, 'month') : [], [state.segments, compRange]);
  const compSeg = useMemo(() => {
    if (!compRange || !compData.length) return null;
    const cl = compData[compData.length - 1];
    return { totalCustomers: cl.totalCustomers, segActive: cl.segActive, segAtRisk: cl.segAtRisk, segLapsed: cl.segLapsed, avgRFMScore: cl.avgRFMScore, migratedAtRiskToActive: cl.migratedAtRiskToActive };
  }, [compRange, compData]);

  const pctData = data.map(m => {
    const total = m.totalCustomers || 1;
    return {
      month: m.month.slice(2),
      New: (m.segNew / total) * 100,
      Active: (m.segActive / total) * 100,
      'At-Risk': (m.segAtRisk / total) * 100,
      Lapsed: (m.segLapsed / total) * 100,
    };
  });

  const revenueData = data.map(m => ({
    month: m.month.slice(2),
    New: m.segNewRevenue,
    Active: m.segActiveRevenue,
    'At-Risk': m.segAtRiskRevenue,
    Lapsed: m.segLapsedRevenue,
  }));

  const recoveryData = data.map(m => ({
    month: m.month.slice(2),
    saved: m.migratedAtRiskToActive,
    lost: -m.migratedActiveToAtRisk,
    net: m.migratedAtRiskToActive - m.migratedActiveToAtRisk,
    reactivated: m.reactivatedFromLapsed,
  }));

  const rfmData = data.map(m => ({ month: m.month.slice(2), rfm: m.avgRFMScore }));

  const migrationMatrix = latest ? {
    labels: ['New', 'Active', 'At-Risk', 'Lapsed'],
    data: [
      [0, latest.segNew * 0.7, latest.segNew * 0.2, latest.segNew * 0.1],
      [0, 0, latest.migratedActiveToAtRisk, 0],
      [0, latest.migratedAtRiskToActive, 0, Math.round(latest.segAtRisk * 0.05)],
      [0, latest.reactivatedFromLapsed, Math.round(latest.segLapsed * 0.02), 0],
    ],
  } : null;

  const segKeys = ['New', 'Active', 'At-Risk', 'Lapsed'];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <TimePeriodToggle tab="segments" tabPeriods={state.tabPeriods} dispatch={dispatch} options={['monthly']} />
      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label="Total Customers" value={latest?.totalCustomers} format="number" status="good" compDelta={compSeg ? pctChange(latest?.totalCustomers, compSeg.totalCustomers) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Active Customers" value={latest?.segActive} format="number" status="good" sparkData={data} sparkKey="segActive" compDelta={compSeg ? pctChange(latest?.segActive, compSeg.segActive) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="At-Risk Customers" value={latest?.segAtRisk} format="number" status={latest?.segAtRisk > 1200 ? 'warning' : 'good'} sparkData={data} sparkKey="segAtRisk" compDelta={compSeg ? pctChange(latest?.segAtRisk, compSeg.segAtRisk) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Lapsed Customers" value={latest?.segLapsed} format="number" status="neutral" compDelta={compSeg ? pctChange(latest?.segLapsed, compSeg.segLapsed) : null} compLabel={compLabel} presentationMode={presentationMode} />
        <KPICard label="Avg RFM Score" value={latest?.avgRFMScore} format="text" status={latest?.avgRFMScore >= 3.0 ? 'good' : 'warning'} presentationMode={presentationMode} />
        <KPICard label="Rescued from At-Risk" value={latest?.migratedAtRiskToActive} format="number" status="good" delta={!compSeg ? calcDelta(data, 'migratedAtRiskToActive') : null} compDelta={compSeg ? pctChange(latest?.migratedAtRiskToActive, compSeg.migratedAtRiskToActive) : null} compLabel={compLabel} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Lifecycle Segment Distribution" tooltip="Monthly breakdown of customers across New, Active, At-Risk, and Lapsed lifecycle segments." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <AreaChart data={pctData} stackOffset="expand">
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${(v*100).toFixed(0)}%`} />
            <Tooltip content={<ChartTooltip formatter={(v) => `${v.toFixed(1)}%`} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {segKeys.map(k => (
              <Area key={k} type="monotone" dataKey={k} stackId="1" fill={LIFECYCLE_COLORS[k]} stroke={LIFECYCLE_COLORS[k]} fillOpacity={0.7} />
            ))}
          </AreaChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Segment Revenue Contribution" tooltip="Revenue generated by each lifecycle segment per month." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={revenueData}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
              <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              {segKeys.map(k => (
                <Bar key={k} dataKey={k} stackId="a" fill={LIFECYCLE_COLORS[k]} />
              ))}
            </BarChart>
          </ResponsiveContainer>
        </div>

        {migrationMatrix && (
          <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
            <ChartHeader title="Migration Matrix (Latest Month)" tooltip="Sankey-style view of customer movements between segments: Active↔At-Risk, New→Active, Lapsed→Reactivated." />
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr>
                    <th style={{ padding: '8px', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11, textAlign: 'left' }}>From ↓ / To →</th>
                    {migrationMatrix.labels.map(l => (
                      <th key={l} style={{ padding: '8px', textAlign: 'center', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{l}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {migrationMatrix.labels.map((from, ri) => (
                    <tr key={from} style={{ borderBottom: `1px solid ${C.divider}` }}>
                      <td style={{ padding: '8px', fontWeight: 600, color: LIFECYCLE_COLORS[from], fontSize: 12 }}>{from}</td>
                      {migrationMatrix.data[ri].map((val, ci) => {
                        const isPositive = (from === 'At-Risk' && migrationMatrix.labels[ci] === 'Active') || (from === 'Lapsed' && migrationMatrix.labels[ci] === 'Active');
                        const isNegative = (from === 'Active' && migrationMatrix.labels[ci] === 'At-Risk');
                        return (
                          <td key={ci} style={{ padding: '8px', textAlign: 'center', background: val === 0 ? 'transparent' : isPositive ? '#E8F5F0' : isNegative ? '#FDE8E8' : '#F6EDDA', fontWeight: val > 0 ? 600 : 400, color: val === 0 ? C.textTertiary : isPositive ? C.success : isNegative ? C.danger : C.textPrimary, borderRadius: 4 }}>
                            {val === 0 ? '—' : formatNumber(Math.round(val))}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="At-Risk Recovery Trend" tooltip="Monthly count of customers rescued from At-Risk back to Active (saved) vs those lost from Active to At-Risk." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={recoveryData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="saved" fill={C.success} name="Saved (At-Risk → Active)" radius={[4,4,0,0]} />
            <Bar dataKey="lost" fill={C.danger} name="Lost (Active → At-Risk)" radius={[4,4,0,0]} />
            <Line type="monotone" dataKey="net" stroke={C.primary} strokeWidth={2} dot={{ r: 4 }} name="Net Recovery" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="RFM Score Trend" tooltip="Average RFM (Recency, Frequency, Monetary) score across the customer base over time. Higher = healthier." />
        <ResponsiveContainer width="100%" height={250}>
          <LineChart data={rfmData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} domain={[2, 4]} />
            <Tooltip content={<ChartTooltip formatter={(v) => v.toFixed(1)} />} />
            <ReferenceLine y={3.0} stroke={C.warning} strokeDasharray="5 5" label={{ value: 'Target', position: 'right', fill: C.warning, fontSize: 11 }} />
            <Line type="monotone" dataKey="rfm" stroke={C.primary} strokeWidth={2} dot={{ r: 4 }} name="Avg RFM Score" />
          </LineChart>
        </ResponsiveContainer>
      </div>

      {/* ─── Subscription Churn Analysis ─── */}
      {(() => {
        const fSubs = filterByDateRange(state.subscriptions || [], start, end, 'month');
        const latestSub = fSubs.length > 0 ? fSubs[fSubs.length - 1] : null;
        const fProductChurn = filterByDateRange(state.productChurn || [], start, end, 'month');
        const productNames = [...new Set(fProductChurn.map(r => r.product))].sort();
        const latestMonth = fProductChurn.length > 0 ? [...new Set(fProductChurn.map(r => r.month))].sort().pop() : null;
        const latestProductData = latestMonth ? fProductChurn.filter(r => r.month === latestMonth) : [];
        const PRODUCT_COLORS = ['#124A2B', '#3B82F6', '#F59E0B', '#D81F26', '#7C3AED'];
        const churnColor = (v) => v < 6 ? '#18917B' : v < 9 ? '#F59E0B' : '#D81F26';

        // Overall churn trend
        const overallChurnTrend = fSubs.map(m => ({
          month: m.month.slice(2),
          voluntary: m.voluntaryChurn,
          involuntary: m.involuntaryChurn,
          total: m.churnRate,
        }));

        // Per-product churn trend
        const churnMonths = [...new Set(fProductChurn.map(r => r.month))].sort();
        const productChurnTrend = churnMonths.map(m => {
          const entry = { month: m.slice(2) };
          productNames.forEach(p => {
            const row = fProductChurn.find(r => r.month === m && r.product === p);
            entry[p] = row ? row.churnRate : null;
          });
          return entry;
        });

        return (
          <>
            <div style={{ marginTop: 4 }}>
              <h3 style={{ margin: '0 0 12px', fontSize: 15, fontWeight: 700, color: C.textPrimary, borderBottom: `2px solid ${C.primary}`, paddingBottom: 8, display: 'inline-block' }}>Subscription Churn Analysis</h3>
            </div>

            {/* Overall Churn KPIs */}
            <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
              <KPICard label="Churn Rate" value={latestSub?.churnRate} format="percent" status={latestSub?.churnRate > 8 ? 'warning' : 'good'} sparkData={fSubs} sparkKey="churnRate" presentationMode={presentationMode} />
              <KPICard label="Voluntary Churn" value={latestSub?.voluntaryChurn} format="percent" status={latestSub?.voluntaryChurn > 6 ? 'warning' : 'good'} presentationMode={presentationMode} />
              <KPICard label="Involuntary Churn" value={latestSub?.involuntaryChurn} format="percent" status={latestSub?.involuntaryChurn > 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
              <KPICard label="Churned Subscribers" value={latestSub?.churnedSubscribers} format="number" status={latestSub?.churnedSubscribers > 400 ? 'warning' : 'good'} sparkData={fSubs} sparkKey="churnedSubscribers" presentationMode={presentationMode} />
            </div>

            {/* Overall Churn Trend */}
            <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
              <ChartHeader title="Overall Churn Rate Trend" tooltip="Monthly churn rate split by voluntary (customer-initiated) and involuntary (payment failure) churn." />
              <ResponsiveContainer width="100%" height={chartHeight}>
                <ComposedChart data={overallChurnTrend}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
                  <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
                  <Tooltip content={<ChartTooltip formatter={(v) => `${v.toFixed(1)}%`} />} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Area type="monotone" dataKey="voluntary" stackId="churn" fill="#F59E0B" stroke="#F59E0B" fillOpacity={0.5} name="Voluntary" />
                  <Area type="monotone" dataKey="involuntary" stackId="churn" fill="#D81F26" stroke="#D81F26" fillOpacity={0.5} name="Involuntary" />
                  <Line type="monotone" dataKey="total" stroke={C.textPrimary} strokeWidth={2} dot={{ r: 4 }} name="Total Churn Rate" />
                </ComposedChart>
              </ResponsiveContainer>
            </div>

            {/* Per-Product Churn Cards */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
              {latestProductData.sort((a, b) => a.churnRate - b.churnRate).map((p, i) => (
                <div key={p.product} style={{ background: C.cardBg, borderRadius: 6, padding: 14, border: `1px solid ${C.cardBorder}`, borderLeft: `4px solid ${churnColor(p.churnRate)}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.textPrimary, marginBottom: 6 }}>{p.product}</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: churnColor(p.churnRate) }}>{p.churnRate.toFixed(1)}%</div>
                  <div style={{ fontSize: 10, color: C.textTertiary, marginTop: 2 }}>
                    Vol {p.voluntaryChurn.toFixed(1)}% · Invol {p.involuntaryChurn.toFixed(1)}%
                  </div>
                  <div style={{ fontSize: 10, color: C.textSecondary, marginTop: 4 }}>
                    {formatNumber(p.activeSubscribers)} active · {formatNumber(p.churnedSubscribers)} churned
                  </div>
                </div>
              ))}
            </div>

            {/* Per-Product Churn Trend */}
            <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
              <ChartHeader title="Churn Rate by Product" tooltip="Monthly churn rate per product. Lower is better. Compare product stickiness over time." />
              <ResponsiveContainer width="100%" height={chartHeight}>
                <LineChart data={productChurnTrend}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.textTertiary }} />
                  <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
                  <Tooltip content={<ChartTooltip formatter={(v) => `${v.toFixed(1)}%`} />} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  {productNames.map((p, i) => (
                    <Line key={p} type="monotone" dataKey={p} stroke={PRODUCT_COLORS[i % PRODUCT_COLORS.length]} strokeWidth={2} dot={{ r: 3 }} name={p} connectNulls />
                  ))}
                </LineChart>
              </ResponsiveContainer>
            </div>

            {/* Per-Product Churn Table */}
            <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
              <ChartHeader title="Product Churn Detail" tooltip="Detailed churn breakdown per product for the latest month." />
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr>{['Product','Active Subs','Churned','Churn Rate','Voluntary','Involuntary','New Subs','Reactivated'].map(h => (
                      <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Product' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {latestProductData.sort((a, b) => a.churnRate - b.churnRate).map((r, i) => (
                      <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                        <td style={{ padding: '8px 10px', fontWeight: 600, color: C.textPrimary }}>{r.product}</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.activeSubscribers)}</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.churnedSubscribers)}</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                          <span style={{ background: churnColor(r.churnRate), color: '#fff', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>{r.churnRate.toFixed(1)}%</span>
                        </td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{r.voluntaryChurn.toFixed(1)}%</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{r.involuntaryChurn.toFixed(1)}%</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.success, fontWeight: 600 }}>+{formatNumber(r.newSubscribers)}</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', color: C.info }}>{formatNumber(r.reactivated)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        );
      })()}

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <ChartHeader title="Segment & List Links" tooltip="Quick links to your Klaviyo segments and lists for easy access." />
          <button onClick={() => { setShowSegLinkForm(!showSegLinkForm); setEditSegLink(null); setSegLinkForm({ name: '', type: 'segment', klaviyo_url: '', description: '', member_count: '' }); }} style={{ padding: '6px 16px', borderRadius: 4, border: 'none', background: showSegLinkForm ? C.textTertiary : C.primary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            {showSegLinkForm ? 'Cancel' : '+ Add Segment'}
          </button>
        </div>

        {showSegLinkForm && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 16, padding: 12, background: C.pageBg, borderRadius: 4 }}>
            <input value={segLinkForm.name} onChange={e => setSegLinkForm({ ...segLinkForm, name: e.target.value })} placeholder="Segment / List name" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary }} />
            <select value={segLinkForm.type} onChange={e => setSegLinkForm({ ...segLinkForm, type: e.target.value })} style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary }}>
              <option value="segment">Segment</option><option value="list">List</option>
            </select>
            <input value={segLinkForm.klaviyo_url} onChange={e => setSegLinkForm({ ...segLinkForm, klaviyo_url: e.target.value })} placeholder="Klaviyo URL" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary, gridColumn: '1/3' }} />
            <input value={segLinkForm.description} onChange={e => setSegLinkForm({ ...segLinkForm, description: e.target.value })} placeholder="Description (optional)" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary }} />
            <input type="number" value={segLinkForm.member_count} onChange={e => setSegLinkForm({ ...segLinkForm, member_count: e.target.value })} placeholder="Member count" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.cardBg, color: C.textPrimary }} />
            <button onClick={saveSegmentLink} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>{editSegLink ? 'Update' : 'Save'}</button>
          </div>
        )}

        {state.segmentLinks.length === 0 && !showSegLinkForm ? (
          <p style={{ fontSize: 13, color: C.textTertiary, textAlign: 'center', padding: 20 }}>No segment links yet. Click "+ Add Segment" to link Klaviyo segments.</p>
        ) : state.segmentLinks.length > 0 && (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr>{['Name', 'Type', 'Members', 'Description', 'Klaviyo', 'Actions'].map(h => (
                  <th key={h} style={{ padding: '8px 10px', textAlign: 'left', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {state.segmentLinks.map(link => (
                  <tr key={link.id} style={{ borderBottom: `1px solid ${C.divider}` }}>
                    <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{link.name}</td>
                    <td style={{ padding: '8px 10px' }}>
                      <span style={{ padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 600, color: '#fff', background: SEGMENT_TYPE_COLORS[link.type] || '#7C3AED' }}>{link.type}</span>
                    </td>
                    <td style={{ padding: '8px 10px', color: C.textSecondary }}>{link.member_count ? formatNumber(link.member_count) : '—'}</td>
                    <td style={{ padding: '8px 10px', color: C.textSecondary, maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{link.description || '—'}</td>
                    <td style={{ padding: '8px 10px' }}>
                      {link.klaviyo_url ? (
                        <a href={link.klaviyo_url} target="_blank" rel="noopener noreferrer" style={{ color: '#7C3AED', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Open &#8599;</a>
                      ) : '—'}
                    </td>
                    <td style={{ padding: '8px 10px' }}>
                      <button onClick={() => { setEditSegLink(link); setSegLinkForm({ name: link.name, type: link.type, klaviyo_url: link.klaviyo_url || '', description: link.description || '', member_count: link.member_count || '' }); setShowSegLinkForm(true); }} style={{ background: 'none', border: 'none', color: C.textTertiary, cursor: 'pointer', fontSize: 11, marginRight: 6 }}>Edit</button>
                      <button onClick={() => deleteSegmentLink(link.id)} style={{ background: 'none', border: 'none', color: '#D81F26', cursor: 'pointer', fontSize: 11 }}>Delete</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function OutreachSection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const { start, end } = state.dateRange;
  const filteredOutreach = useMemo(() => filterByDateRange(state.outreach, start, end, 'week'), [state.outreach, start, end]);
  const period = state.tabPeriods.outreach || 'weekly';
  const data = period === 'monthly' ? aggregateOutreachByMonth(filteredOutreach) : filteredOutreach;
  const weeks = [...new Set(data.map(r => r.week))].sort();
  const latestWeek = weeks[weeks.length - 1];
  const latestData = data.filter(r => r.week === latestWeek);
  const latestRev = latestData.reduce((s, r) => s + (r.revenue || 0), 0);
  const latestCost = latestData.reduce((s, r) => s + (r.cost || 0), 0);
  const outreachROAS = latestCost > 0 ? latestRev / latestCost : 0;
  const latestWA = latestData.find(r => r.channel === 'WhatsApp');
  const latestSMS = latestData.find(r => r.channel === 'SMS Blast');
  const latestCall = latestData.find(r => r.channel === 'Personal Call');

  const channels = [...new Set(data.map(r => r.channel))];

  const weeklyRevData = weeks.map(w => {
    const row = { week: w.slice(5) };
    data.filter(r => r.week === w).forEach(r => { row[r.channel] = r.revenue; });
    return row;
  });

  const last4 = weeks.slice(-4);
  const channelSummary = channels.map(ch => {
    const rows = data.filter(r => r.channel === ch && last4.includes(r.week));
    const totalSends = rows.reduce((s, r) => s + r.sends, 0);
    const totalResponses = rows.reduce((s, r) => s + r.responses, 0);
    const totalConv = rows.reduce((s, r) => s + r.conversions, 0);
    const totalRev = rows.reduce((s, r) => s + r.revenue, 0);
    const totalCost = rows.reduce((s, r) => s + r.cost, 0);
    return {
      channel: ch,
      sends: totalSends,
      responseRate: totalSends > 0 ? (totalResponses / totalSends) * 100 : 0,
      conversions: totalConv,
      conversionRate: totalSends > 0 ? (totalConv / totalSends) * 100 : 0,
      revenue: totalRev,
      cost: totalCost,
      roas: totalCost > 0 ? totalRev / totalCost : 0,
      costPerConversion: totalConv > 0 ? totalCost / totalConv : 0,
    };
  }).sort((a, b) => b.revenue - a.revenue);

  const responseRateData = channelSummary.map(c => ({
    name: c.channel,
    responseRate: c.responseRate,
    fill: CRM_CHANNEL_COLORS[c.channel] || C.info,
  }));

  const costPerConvData = channelSummary.filter(c => c.costPerConversion > 0).map(c => ({
    name: c.channel,
    costPerConversion: c.costPerConversion,
    fill: CRM_CHANNEL_COLORS[c.channel] || C.info,
  }));

  const waData = weeks.map(w => {
    const r = data.find(d => d.week === w && d.channel === 'WhatsApp');
    return r ? { week: w.slice(5), sends: r.sends, responseRate: r.responseRate, conversionRate: r.conversionRate } : null;
  }).filter(Boolean);

  const outPeriodLabel = period === 'monthly' ? 'Month' : 'Week';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <TimePeriodToggle tab="outreach" tabPeriods={state.tabPeriods} dispatch={dispatch} />
      <div className="crm-kpi-grid" style={{ display: 'grid', gap: 12 }}>
        <KPICard label={`Outreach Revenue (${outPeriodLabel})`} value={latestRev} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="WhatsApp Response Rate" value={latestWA?.responseRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="SMS Conversion Rate" value={latestSMS?.conversionRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="Personal Call Conv. Rate" value={latestCall?.conversionRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="Outreach ROAS" value={outreachROAS} format="multiplier" status={outreachROAS < 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Outreach Revenue by Channel" tooltip="Weekly revenue from direct outreach channels: WhatsApp, SMS, Push, and Direct Mail." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <BarChart data={weeklyRevData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {channels.map(ch => (
              <Bar key={ch} dataKey={ch} stackId="a" fill={CRM_CHANNEL_COLORS[ch] || C.info} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="crm-2col-grid" style={{ display: 'grid', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Response Rate by Channel" tooltip="Response rates for each outreach channel, showing customer engagement levels." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={responseRateData} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
              <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 11, fill: C.textSecondary }} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatPercent(v)} />} />
              <Bar dataKey="responseRate" radius={[0,4,4,0]}>
                {responseRateData.map((e, i) => <Cell key={i} fill={e.fill} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Cost per Conversion by Channel" tooltip="Marketing cost divided by conversions for each outreach channel — lower is better." />
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={costPerConvData} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
              <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${v.toFixed(0)}`} />
              <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 11, fill: C.textSecondary }} />
              <Tooltip content={<ChartTooltip formatter={(v) => formatCurrencyDecimal(v)} />} />
              <Bar dataKey="costPerConversion" radius={[0,4,4,0]}>
                {costPerConvData.map((e, i) => <Cell key={i} fill={e.fill} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Channel Performance (Last 4 Weeks)" tooltip="Per-channel breakdown of sends, responses, conversions, and revenue for the most recent 4 weeks." />
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Channel','Sends','Response Rate','Conversions','Conv. Rate','Revenue','Cost','ROAS','Cost/Conv'].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Channel' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {channelSummary.map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                  <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{r.channel}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.sends)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatPercent(r.responseRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.conversions)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatPercent(r.conversionRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(r.revenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.cost)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: r.roas >= 10 ? C.success : r.roas >= 3 ? C.info : C.warning }}>{formatMultiplier(r.roas)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrencyDecimal(r.costPerConversion)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="WhatsApp Engagement Trend" tooltip="Weekly WhatsApp-specific sends, responses, and response rate trend." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={waData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
            <Tooltip content={<ChartTooltip formatter={(v, name) => name.includes('Rate') ? formatPercent(v) : formatNumber(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="sends" fill="#18917B" name="Sends" opacity={0.4} radius={[4,4,0,0]} />
            <Line yAxisId="right" type="monotone" dataKey="responseRate" stroke="#18917B" strokeWidth={2} dot={{ r: 3 }} name="Response Rate" />
            <Line yAxisId="right" type="monotone" dataKey="conversionRate" stroke={C.primary} strokeWidth={2} dot={{ r: 3 }} name="Conversion Rate" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function IncrementalitySection({ state, presentationMode }) {
  const chartHeight = presentationMode ? 400 : 300;
  const totalIncRev = state.activityROI.reduce((s, r) => s + r.incrementalRevenue, 0);
  const totalCost = state.activityROI.reduce((s, r) => s + r.totalCost, 0);
  const avgROI = totalCost > 0 ? totalIncRev / totalCost : 0;
  const activeTests = state.holdoutTests.filter(t => t.status === 'active').length;
  const highestLift = state.beforeAfter.reduce((max, r) => r.lift > max.lift ? r : max, { lift: 0 });
  const roiColor = (v) => v > 10 ? '#18917B' : v > 3 ? '#2D8B6E' : v > 1 ? '#F59E0B' : '#D81F26';

  const activities = [...new Set(state.beforeAfter.map(r => r.activity))];

  const roiSorted = [...state.activityROI].sort((a, b) => b.incrementalRevenue - a.incrementalRevenue);
  const roiTotalCost = roiSorted.reduce((s, r) => s + r.totalCost, 0);
  const roiTotalAttr = roiSorted.reduce((s, r) => s + r.attributedRevenue, 0);
  const roiTotalInc = roiSorted.reduce((s, r) => s + r.incrementalRevenue, 0);
  const roiTotalCust = roiSorted.reduce((s, r) => s + r.customersInfluenced, 0);
  const roiAvgROI = roiTotalCost > 0 ? roiTotalInc / roiTotalCost : 0;

  const holdoutBarData = [...state.holdoutTests].sort((a, b) => b.incrementalRevenue - a.incrementalRevenue).map(t => ({
    name: t.testName,
    incrementalRevenue: t.incrementalRevenue,
    fill: t.confidence >= 0.90 ? C.success : t.confidence >= 0.80 ? C.info : C.warning,
  }));

  const scatterData = state.activityROI.map(r => ({
    x: r.totalCost,
    y: r.incrementalRevenue,
    z: r.customersInfluenced,
    name: r.activity,
    channel: r.channel,
  }));

  const confBadge = (conf) => {
    const bg = conf >= 0.90 ? '#E8F5F0' : conf >= 0.80 ? '#E0EEE9' : '#FEF3C7';
    const color = conf >= 0.90 ? '#124A2B' : conf >= 0.80 ? '#18917B' : '#92400E';
    return <span style={{ background: bg, color, padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>{(conf * 100).toFixed(0)}%</span>;
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 12 }}>
        <KPICard label="Total Incremental Revenue (Q1)" value={totalIncRev} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="Avg Incremental ROI" value={avgROI} format="multiplier" status={avgROI < 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label="Active Holdout Tests" value={activeTests} format="number" status="good" presentationMode={presentationMode} />
        <KPICard label="Highest Lift Activity" value={`${highestLift.activity?.split(' ').slice(0,2).join(' ')} +${highestLift.lift?.toFixed(0)}%`} format="text" status="good" presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Before / After Analysis" tooltip="Compares key metrics before and after major CRM activities were launched. Shows absolute lift percentage." />
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Activity','Metric','Before','After','Lift',''].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Activity' || h === 'Metric' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {activities.map(act => {
                const rows = state.beforeAfter.filter(r => r.activity === act);
                return rows.map((r, i) => (
                  <tr key={`${act}-${i}`} style={{ borderBottom: `1px solid ${C.divider}` }}>
                    {i === 0 && <td rowSpan={rows.length} style={{ padding: '8px 10px', fontWeight: 600, color: C.textPrimary, verticalAlign: 'top', borderRight: `1px solid ${C.divider}` }}>{act}</td>}
                    <td style={{ padding: '8px 10px', color: C.textSecondary }}>{r.metric}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{r.unit === 'currency' ? formatCurrency(r.beforeValue) : formatPercent(r.beforeValue)}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{r.unit === 'currency' ? formatCurrency(r.afterValue) : formatPercent(r.afterValue)}</td>
                    <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                      <span style={{ background: r.lift > 50 ? '#E8F5F0' : r.lift > 20 ? '#E0EEE9' : '#FEF3C7', color: r.lift > 50 ? '#124A2B' : r.lift > 20 ? '#18917B' : '#92400E', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>+{r.lift.toFixed(1)}%</span>
                    </td>
                    <td style={{ padding: '8px 10px', width: 120 }}>
                      <div style={{ background: C.divider, borderRadius: 4, height: 8, position: 'relative' }}>
                        <div style={{ background: r.lift > 50 ? C.success : r.lift > 20 ? C.info : C.warning, borderRadius: 4, height: 8, width: `${Math.min(r.lift / 2.5, 100)}%` }} />
                      </div>
                    </td>
                  </tr>
                ));
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Control vs Test Group Results" tooltip="Side-by-side holdout test results: control group conversion/revenue vs exposed group, with incremental lift." />
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Test Name','Period','Sample','Control Conv%','Exposed Conv%','Conv Lift','Incremental Rev','Confidence'].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Test Name' || h === 'Period' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {state.holdoutTests.map((t, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                  <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{t.testName}</td>
                  <td style={{ padding: '8px 10px', color: C.textSecondary }}>{t.testPeriod}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(t.sampleSize)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatPercent(t.controlConversionRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatPercent(t.exposedConversionRate)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.success, fontWeight: 600 }}>+{formatPercent(t.incrementalConversionLift)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(t.incrementalRevenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right' }}>{confBadge(t.confidence)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Incremental Revenue by Test" tooltip="Ranked bar chart of incremental revenue generated by each holdout test." />
        <ResponsiveContainer width="100%" height={250}>
          <BarChart data={holdoutBarData} layout="vertical">
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis type="number" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <YAxis type="category" dataKey="name" width={180} tick={{ fontSize: 11, fill: C.textSecondary }} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Bar dataKey="incrementalRevenue" radius={[0,4,4,0]}>
              {holdoutBarData.map((e, i) => <Cell key={i} fill={e.fill} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Activity ROI Summary" tooltip="Cost, attributed revenue, incremental revenue, and ROI multiplier for each CRM activity." />
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr>{['Activity','Channel','Cost','Attributed Rev','Incremental Rev','ROI','Customers'].map(h => (
                <th key={h} style={{ padding: '8px 10px', textAlign: h === 'Activity' || h === 'Channel' ? 'left' : 'right', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11 }}>{h}</th>
              ))}</tr>
            </thead>
            <tbody>
              {roiSorted.map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                  <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{r.activity}</td>
                  <td style={{ padding: '8px 10px', color: C.textSecondary }}>{r.channel}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.totalCost)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatCurrency(r.attributedRevenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: C.textPrimary }}>{formatCurrency(r.incrementalRevenue)}</td>
                  <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                    <span style={{ background: roiColor(r.incrementalROI), color: '#fff', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>{formatMultiplier(r.incrementalROI)}</span>
                  </td>
                  <td style={{ padding: '8px 10px', textAlign: 'right', color: C.textSecondary }}>{formatNumber(r.customersInfluenced)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ borderTop: `2px solid ${C.cardBorder}`, background: C.divider }}>
                <td colSpan={2} style={{ padding: '8px 10px', fontWeight: 700, color: C.textPrimary }}>TOTAL</td>
                <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: C.textPrimary }}>{formatCurrency(roiTotalCost)}</td>
                <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: C.textPrimary }}>{formatCurrency(roiTotalAttr)}</td>
                <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: C.textPrimary }}>{formatCurrency(roiTotalInc)}</td>
                <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                  <span style={{ background: roiColor(roiAvgROI), color: '#fff', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>{formatMultiplier(roiAvgROI)}</span>
                </td>
                <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: C.textPrimary }}>{formatNumber(roiTotalCust)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <ChartHeader title="Cost vs Incremental Revenue" tooltip="Scatter plot of CRM activity cost vs incremental revenue. Bubble size = customers influenced." />
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ScatterChart>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis type="number" dataKey="x" name="Cost" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} label={{ value: 'Total Cost (£)', position: 'bottom', offset: -5, style: { fontSize: 11, fill: C.textSecondary } }} />
            <YAxis type="number" dataKey="y" name="Incremental Revenue" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} label={{ value: 'Incremental Revenue (£)', angle: -90, position: 'insideLeft', style: { fontSize: 11, fill: C.textSecondary } }} />
            <Tooltip content={({ active, payload }) => {
              if (!active || !payload?.length) return null;
              const d = payload[0].payload;
              return (
                <div style={{ background: C.cardBg, border: `1px solid ${C.cardBorder}`, borderRadius: 4, padding: '10px 14px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
                  <p style={{ margin: 0, fontWeight: 600, fontSize: 12, color: C.textPrimary }}>{d.name}</p>
                  <p style={{ margin: '4px 0 0', fontSize: 12, color: C.textSecondary }}>Cost: {formatCurrency(d.x)}</p>
                  <p style={{ margin: '2px 0 0', fontSize: 12, color: C.textSecondary }}>Inc. Rev: {formatCurrency(d.y)}</p>
                  <p style={{ margin: '2px 0 0', fontSize: 12, color: C.textSecondary }}>Customers: {formatNumber(d.z)}</p>
                </div>
              );
            }} />
            <Scatter data={scatterData} fill={C.primary}>
              {scatterData.map((e, i) => (
                <Cell key={i} fill={CRM_CHANNEL_COLORS[e.channel] || C.primary} r={Math.max(6, Math.min(20, Math.sqrt(e.z) / 5))} />
              ))}
            </Scatter>
            <ReferenceLine segment={[{ x: 0, y: 0 }, { x: 15000, y: 15000 }]} stroke={C.textTertiary} strokeDasharray="5 5" label={{ value: '1x ROI', position: 'end', fill: C.textTertiary, fontSize: 10 }} />
          </ScatterChart>
        </ResponsiveContainer>
      </div>

      {state.activityLog && state.activityLog.length > 0 && (
        <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <ChartHeader title="Dashboard Change History" tooltip="Log of all data changes made to the dashboard, including who made them and when." />
          <div style={{ maxHeight: 400, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr>{['Time', 'Action', 'Category', 'Detail', 'User'].map(h => (
                  <th key={h} style={{ padding: '8px 10px', textAlign: 'left', borderBottom: `2px solid ${C.cardBorder}`, color: C.textSecondary, fontWeight: 600, fontSize: 11, position: 'sticky', top: 0, background: C.cardBg }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {state.activityLog.slice(0, 50).map((log, i) => (
                  <tr key={log.id || i} style={{ borderBottom: `1px solid ${C.divider}` }}>
                    <td style={{ padding: '8px 10px', color: C.textTertiary, fontSize: 11, whiteSpace: 'nowrap' }}>{timeAgo(log.timestamp)}</td>
                    <td style={{ padding: '8px 10px', fontWeight: 500, color: C.textPrimary }}>{log.action}</td>
                    <td style={{ padding: '8px 10px' }}>
                      <span style={{ padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 600, color: '#fff', background: CATEGORY_LOG_COLORS[log.category] || C.primary }}>{log.category}</span>
                    </td>
                    <td style={{ padding: '8px 10px', color: C.textSecondary, maxWidth: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{log.detail}</td>
                    <td style={{ padding: '8px 10px', color: C.textSecondary }}>{log.user}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function ChannelCostEntryForm({ dispatch, existingCosts }) {
  const [startMonth, setStartMonth] = useState('');
  const [endMonth, setEndMonth] = useState('');
  const [costGrid, setCostGrid] = useState(null);
  const [gridMonths, setGridMonths] = useState([]);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);

  const generateGrid = () => {
    setError(null); setSuccess(null);
    if (!startMonth || !endMonth) { setError('Please select both start and end month.'); return; }
    if (startMonth > endMonth) { setError('Start month must be before or equal to end month.'); return; }
    const months = [];
    let [y, m] = startMonth.split('-').map(Number);
    const [ey, em] = endMonth.split('-').map(Number);
    while (y < ey || (y === ey && m <= em)) {
      months.push(`${y}-${String(m).padStart(2, '0')}`);
      m++;
      if (m > 12) { m = 1; y++; }
    }
    if (months.length > 24) { setError('Maximum 24 months at a time.'); return; }
    const grid = {};
    months.forEach(mo => {
      grid[mo] = {};
      CHANNEL_DEFS.forEach(ch => {
        const existing = (existingCosts || []).find(c => c.month === mo && c.channel === ch.key);
        grid[mo][ch.key] = existing ? String(existing.cost) : '';
      });
      const existingNote = (existingCosts || []).find(c => c.month === mo);
      grid[mo].notes = existingNote?.notes || '';
    });
    setGridMonths(months);
    setCostGrid(grid);
  };

  const updateCell = (month, field, value) => {
    setCostGrid(prev => ({ ...prev, [month]: { ...prev[month], [field]: value } }));
  };

  const importCosts = () => {
    setError(null); setSuccess(null);
    if (!costGrid || gridMonths.length === 0) { setError('Generate the grid first.'); return; }
    const newRecords = [];
    gridMonths.forEach(mo => {
      CHANNEL_DEFS.forEach(ch => {
        const val = costGrid[mo]?.[ch.key];
        if (val !== '' && val !== undefined && val !== null) {
          newRecords.push({ month: mo, channel: ch.key, cost: Number(val) || 0, notes: costGrid[mo]?.notes || '' });
        }
      });
    });
    if (newRecords.length === 0) { setError('No costs entered. Fill in at least one cell.'); return; }
    // Merge: keep existing costs for months outside the grid, replace within
    const outside = (existingCosts || []).filter(c => !gridMonths.includes(c.month));
    const merged = [...outside, ...newRecords];
    dispatch({ type: 'LOAD_DATA', source: 'channelCosts', payload: merged });
    setSuccess(`Imported ${newRecords.length} cost records for ${gridMonths.length} month(s).`);
  };

  const inputStyle = { padding: '6px 8px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: C.inputBg, color: C.textPrimary, fontSize: 13, width: '100%', boxSizing: 'border-box' };

  return (
    <div style={{ background: C.cardBg, borderRadius: 6, padding: 20, border: `1px solid ${C.cardBorder}` }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 18 }}>💰</span>
        <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Channel Cost Entry</h3>
        <span style={{ fontSize: 12, color: C.textSecondary, marginLeft: 'auto' }}>Enter monthly platform costs for ROI tracking</span>
      </div>

      <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end', flexWrap: 'wrap', marginBottom: 12 }}>
        <div>
          <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Start Month</label>
          <input type="month" value={startMonth} onChange={e => setStartMonth(e.target.value)} style={{ ...inputStyle, width: 160 }} />
        </div>
        <div>
          <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>End Month</label>
          <input type="month" value={endMonth} onChange={e => setEndMonth(e.target.value)} style={{ ...inputStyle, width: 160 }} />
        </div>
        <button onClick={generateGrid} style={{ padding: '7px 20px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>Generate Grid</button>
      </div>

      {error && <div style={{ padding: '8px 12px', background: '#FEF2F2', color: C.danger, borderRadius: 4, fontSize: 13, marginBottom: 12 }}>{error}</div>}
      {success && <div style={{ padding: '8px 12px', background: '#F0FDF4', color: '#166534', borderRadius: 4, fontSize: 13, marginBottom: 12 }}>{success}</div>}

      {costGrid && gridMonths.length > 0 && (
        <>
          <div style={{ overflowX: 'auto', marginBottom: 12 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: `2px solid ${C.cardBorder}` }}>
                  <th style={{ padding: '8px 6px', textAlign: 'left', fontWeight: 700, color: C.textPrimary, whiteSpace: 'nowrap' }}>Month</th>
                  {CHANNEL_DEFS.map(ch => (
                    <th key={ch.key} style={{ padding: '8px 6px', textAlign: 'right', fontWeight: 700, color: ch.color, whiteSpace: 'nowrap', minWidth: 100 }}>{ch.key}</th>
                  ))}
                  <th style={{ padding: '8px 6px', textAlign: 'left', fontWeight: 700, color: C.textSecondary, minWidth: 140 }}>Notes</th>
                </tr>
              </thead>
              <tbody>
                {gridMonths.map(mo => (
                  <tr key={mo} style={{ borderBottom: `1px solid ${C.cardBorder}` }}>
                    <td style={{ padding: '6px', fontWeight: 600, color: C.textPrimary, whiteSpace: 'nowrap' }}>{mo}</td>
                    {CHANNEL_DEFS.map(ch => (
                      <td key={ch.key} style={{ padding: '4px 6px' }}>
                        <input type="number" min="0" step="100" placeholder="0" value={costGrid[mo]?.[ch.key] || ''} onChange={e => updateCell(mo, ch.key, e.target.value)} style={{ ...inputStyle, textAlign: 'right', width: 100 }} />
                      </td>
                    ))}
                    <td style={{ padding: '4px 6px' }}>
                      <input type="text" placeholder="Optional notes" value={costGrid[mo]?.notes || ''} onChange={e => updateCell(mo, 'notes', e.target.value)} style={{ ...inputStyle, minWidth: 120 }} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <button onClick={importCosts} style={{ padding: '8px 24px', borderRadius: 4, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Import Costs</button>
            <span style={{ fontSize: 12, color: C.textSecondary }}>Costs for months in this range will be updated. Other months remain unchanged.</span>
          </div>
        </>
      )}
    </div>
  );
}

function DataImportSection({ state, dispatch, onOpenSettings, currentUser }) {
  const connStr = getNeonConnection();
  const username = currentUser?.displayName || 'Anonymous';
  const [importLogs, setImportLogs] = useState([]);

  const loadLogs = useCallback(async () => {
    if (connStr) {
      try {
        const rows = await neonQuery(connStr, 'SELECT * FROM import_log ORDER BY created_at DESC LIMIT 50');
        setImportLogs(rows);
      } catch (_) { /* fall back to localStorage */ setImportLogs(JSON.parse(localStorage.getItem('crm_import_log') || '[]')); }
    } else {
      setImportLogs(JSON.parse(localStorage.getItem('crm_import_log') || '[]'));
    }
  }, [connStr]);

  useEffect(() => { loadLogs(); }, [loadLogs]);

  const handleLogImport = useCallback(async ({ dataset, inputMode, rowCount, summary }) => {
    const entry = { dataset, input_mode: inputMode, row_count: rowCount, summary, imported_by: username, created_at: new Date().toISOString() };
    if (connStr) {
      try {
        await neonQuery(connStr, 'INSERT INTO import_log (dataset, input_mode, row_count, summary, imported_by) VALUES ($1, $2, $3, $4, $5)', [dataset, inputMode, rowCount, summary, username]);
      } catch (_) { /* fall through to localStorage */ }
    }
    // Always save to localStorage as backup
    const local = JSON.parse(localStorage.getItem('crm_import_log') || '[]');
    local.unshift(entry);
    localStorage.setItem('crm_import_log', JSON.stringify(local.slice(0, 100)));
    loadLogs();
  }, [connStr, username, loadLogs]);

  const clearLogs = useCallback(async () => {
    if (connStr) { try { await neonQuery(connStr, 'DELETE FROM import_log'); } catch (_) {} }
    localStorage.removeItem('crm_import_log');
    setImportLogs([]);
  }, [connStr]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <AIDataImporter dispatch={dispatch} onOpenSettings={onOpenSettings} onLogImport={handleLogImport} dashboardState={state} />
      <ChannelCostEntryForm dispatch={dispatch} existingCosts={state.channelCosts} />
      <ImportActivityLog logs={importLogs} onClear={clearLogs} currentUser={currentUser} />
      <h3 style={{ margin: '12px 0 0', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Manual CSV Upload</h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: 16 }}>
        <CSVUploader label="Email & Flows Data" source="emailFlows" requiredHeaders={['week','type','sends','revenue']} dispatch={dispatch} />
        <CSVUploader label="Milestone Reward Data" source="loyalty" requiredHeaders={['month','totalMembers','memberAOV']} dispatch={dispatch} />
        <CSVUploader label="Customer Segments Data" source="segments" requiredHeaders={['month','segNew','segActive','segAtRisk','segLapsed']} dispatch={dispatch} />
        <CSVUploader label="Before/After Analysis" source="beforeAfter" requiredHeaders={['activity','metric','beforeValue','afterValue','lift']} dispatch={dispatch} />
        <CSVUploader label="Holdout Test Results" source="holdoutTests" requiredHeaders={['testName','controlConversionRate','exposedConversionRate','incrementalRevenue']} dispatch={dispatch} />
        <CSVUploader label="Activity ROI Data" source="activityROI" requiredHeaders={['activity','totalCost','incrementalRevenue','incrementalROI']} dispatch={dispatch} />
        <CSVUploader label="Revenue Data" source="revenue" requiredHeaders={['week','totalRevenue','netRevenue']} dispatch={dispatch} />
        <CSVUploader label="Subscription Data" source="subscriptions" requiredHeaders={['month','activeSubscribers','mrr','churnRate']} dispatch={dispatch} />
      </div>
      <div style={{ display: 'flex', gap: 12, justifyContent: 'center', paddingTop: 12 }}>
        <button onClick={() => dispatch({ type: 'CLEAR_ALL' })} style={{ padding: '10px 24px', borderRadius: 4, border: `1px solid ${C.danger}`, background: 'transparent', color: C.danger, fontWeight: 600, cursor: 'pointer', fontSize: 13 }}>Clear All Data</button>
      </div>
    </div>
  );
}

// ─── MAIN APP ───
export default function App() {
  const [state, dispatch] = useReducer(reducer, initialState);
  const [presentationMode, setPresentationMode] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [userMgmtOpen, setUserMgmtOpen] = useState(false);
  const [currentUser, setCurrentUser] = useState(() => {
    const id = localStorage.getItem('crm_user_id');
    if (!id) return null;
    return { id: Number(id), username: localStorage.getItem('crm_username') || '', displayName: localStorage.getItem('crm_display_name') || '', role: localStorage.getItem('crm_user_role') || 'user' };
  });

  const openSettings = useCallback(() => setSettingsOpen(true), []);

  // Logged dispatch: wraps dispatch to auto-log tracked actions
  const loggedDispatch = useCallback((action) => {
    dispatch(action);
    const user = currentUser?.displayName || 'Anonymous';
    if (action.type === 'LOAD_DATA') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Data Import', category: action.source, detail: `Loaded ${Array.isArray(action.payload) ? action.payload.length : 0} rows into ${action.source}`, user } });
    }
    if (action.type === 'APPEND_DATA') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Flow Added', category: action.source || 'emailFlows', detail: `Added ${action.payload?.length || 0} row(s)`, user } });
    }
    if (action.type === 'ADD_SEGMENT_LINK') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Segment Link Created', category: 'segments', detail: `Created: ${action.payload?.name || ''}`, user } });
    }
    if (action.type === 'DELETE_SEGMENT_LINK') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Segment Link Deleted', category: 'segments', detail: `Deleted segment link`, user } });
    }
    if (action.type === 'RESET_DEMO') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Reset to Demo', category: 'system', detail: 'All data reset to demo values', user } });
    }
    if (action.type === 'CLEAR_ALL') {
      dispatch({ type: 'LOG_ACTIVITY', payload: { action: 'Clear All Data', category: 'system', detail: 'All data cleared', user } });
    }
  }, [currentUser]);

  // Persist activity log
  useEffect(() => {
    if (state.activityLog.length > 0) {
      localStorage.setItem('crm_activity_log', JSON.stringify(state.activityLog.slice(0, 200)));
    }
  }, [state.activityLog]);

  // Load activity log on mount
  useEffect(() => {
    const saved = localStorage.getItem('crm_activity_log');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) dispatch({ type: 'SET_ACTIVITY_LOG', payload: parsed });
      } catch (_) {}
    }
  }, []);

  // ── Load dashboard data from Neon on mount, then persist changes ──
  const neonInitDone = useRef(false);
  const skipPersist = useRef(true); // skip first persist cycle (initial load)
  useEffect(() => {
    const conn = getNeonConnection();
    if (!conn || neonInitDone.current) return;
    neonInitDone.current = true;
    loadAllDashboardData(conn).then(data => {
      if (data) {
        DATA_KEYS.forEach(key => {
          if (data[key] && Array.isArray(data[key]) && data[key].length > 0) {
            dispatch({ type: 'LOAD_DATA', source: key, payload: data[key] });
          }
        });
      }
      // Allow persisting after initial load completes
      setTimeout(() => { skipPersist.current = false; }, 500);
    }).catch(() => { skipPersist.current = false; });
  }, []);

  // ── Persist dashboard data to Neon when it changes ──
  const prevDataRef = useRef({});
  useEffect(() => {
    if (skipPersist.current) return;
    const conn = getNeonConnection();
    if (!conn) return;
    DATA_KEYS.forEach(key => {
      if (state[key] !== prevDataRef.current[key] && state[key] !== undefined) {
        saveDashboardData(conn, key, state[key]);
      }
    });
    const refs = {};
    DATA_KEYS.forEach(key => { refs[key] = state[key]; });
    prevDataRef.current = refs;
  }, [state.emailFlows, state.loyalty, state.segments, state.outreach, state.beforeAfter, state.holdoutTests, state.activityROI, state.revenue, state.subscriptions, state.milestoneProducts, state.whatsappFlows, state.postcardFlows, state.channelCosts, state.productChurn]);

  const handleLogout = useCallback(() => {
    localStorage.removeItem('crm_user_id');
    localStorage.removeItem('crm_username');
    localStorage.removeItem('crm_display_name');
    localStorage.removeItem('crm_user_role');
    setCurrentUser(null);
  }, []);

  if (!currentUser) {
    return <LoginScreen onLogin={setCurrentUser} />;
  }

  return (
    <div style={{ minHeight: '100vh', background: C.pageBg, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }}>
      <ResponsiveStyles />
      <SettingsModal open={settingsOpen} onClose={() => setSettingsOpen(false)} currentUser={currentUser} />
      <UserManagementModal open={userMgmtOpen} onClose={() => setUserMgmtOpen(false)} currentUser={currentUser} />
      <header className="crm-header" style={{ background: C.cardBg, borderBottom: `1px solid ${C.cardBorder}`, padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div className="crm-header-left" style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <OmniPetLogo height={28} />
          <div>
            <h1 style={{ margin: 0, fontSize: presentationMode ? 24 : 20, fontWeight: 700, color: C.textPrimary }}>CRM Dashboard</h1>
            <p style={{ margin: '2px 0 0', fontSize: 13, color: C.textSecondary }}>Subscription DTC — CRM Performance & Incrementality</p>
          </div>
        </div>
        <div className="crm-header-right" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 10px', borderRadius: 4, background: C.divider }}>
            <div style={{ width: 24, height: 24, borderRadius: '50%', background: C.primary, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700 }}>
              {currentUser.displayName[0]?.toUpperCase() || '?'}
            </div>
            <span style={{ fontSize: 12, fontWeight: 600, color: C.textPrimary }}>{currentUser.displayName}</span>
            <button onClick={handleLogout} style={{ fontSize: 11, color: C.textTertiary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline', padding: 0 }}>Log out</button>
          </div>
          {currentUser.role === 'admin' && (
            <button onClick={() => setUserMgmtOpen(true)} title="User Management" style={{ padding: '6px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
              <span style={{ fontSize: 15 }}>&#128101;</span> Users
            </button>
          )}
          <button onClick={openSettings} title="Settings" style={{ padding: '8px 12px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 16, cursor: 'pointer', lineHeight: 1 }}>&#9881;</button>
          <button className="crm-pres-btn" onClick={() => setPresentationMode(!presentationMode)} style={{ padding: '8px 16px', borderRadius: 4, border: `1px solid ${C.cardBorder}`, background: presentationMode ? C.primary : 'transparent', color: presentationMode ? '#fff' : C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            {presentationMode ? 'Exit Presentation' : 'Presentation Mode'}
          </button>
        </div>
      </header>

      <div className="crm-main" style={{ maxWidth: 1400, margin: '0 auto', padding: '0 24px 40px' }}>
        <TabNav tabs={TABS} active={state.activeTab} onSelect={t => dispatch({ type: 'SET_TAB', payload: t })} />
        {['overview', 'email', 'loyalty', 'segments'].includes(state.activeTab) && (
          <div style={{ paddingTop: 12 }}>
            <DateRangePicker state={state} dispatch={dispatch} />
          </div>
        )}
        <div style={{ paddingTop: 20 }}>
          {state.activeTab === 'overview' && <OverviewSection state={state} presentationMode={presentationMode} dispatch={dispatch} />}
          {state.activeTab === 'email' && <EmailFlowsSection state={state} presentationMode={presentationMode} dispatch={loggedDispatch} />}
          {state.activeTab === 'whatsapp' && <WhatsAppSection state={state} presentationMode={presentationMode} dispatch={loggedDispatch} />}
          {state.activeTab === 'postcard' && <PostcardSection state={state} presentationMode={presentationMode} dispatch={loggedDispatch} />}
          {state.activeTab === 'loyalty' && <LoyaltySection state={state} presentationMode={presentationMode} dispatch={loggedDispatch} />}
          {state.activeTab === 'segments' && <SegmentsSection state={state} presentationMode={presentationMode} dispatch={loggedDispatch} />}
          {state.activeTab === 'incrementality' && <IncrementalitySection state={state} presentationMode={presentationMode} />}
          {state.activeTab === 'initiatives' && <InitiativesSection onOpenSettings={openSettings} currentUser={currentUser} presentationMode={presentationMode} onLogActivity={(entry) => dispatch({ type: 'LOG_ACTIVITY', payload: entry })} activityLog={state.activityLog} dashboardState={state} dispatch={loggedDispatch} />}
          {state.activeTab === 'import' && <DataImportSection state={state} dispatch={loggedDispatch} onOpenSettings={openSettings} currentUser={currentUser} />}
        </div>
      </div>
    </div>
  );
}
