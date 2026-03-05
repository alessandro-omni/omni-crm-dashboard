import React, { useReducer, useMemo, useCallback, useRef, useState, useEffect } from 'react';
import Papa from 'papaparse';
import { neon } from '@neondatabase/serverless';
import {
  LineChart, Line, AreaChart, Area, BarChart, Bar,
  ComposedChart, PieChart, Pie, Cell, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine
} from 'recharts';

// ─── COLOR TOKENS ───
const C = {
  primary: '#4F46E5',
  success: '#10B981',
  danger: '#EF4444',
  warning: '#F59E0B',
  secondary: '#8B5CF6',
  info: '#3B82F6',
  textPrimary: '#111827',
  textSecondary: '#6B7280',
  textTertiary: '#9CA3AF',
  cardBg: '#FFFFFF',
  pageBg: '#F9FAFB',
  cardBorder: '#E5E7EB',
  divider: '#F3F4F6',
};

const CRM_CHANNEL_COLORS = {
  'Email Campaign': '#7C3AED',
  'Welcome Series': '#3B82F6',
  'Win-Back 60d': '#F59E0B',
  'Abandoned Cart': '#EF4444',
  'Post-Purchase Upsell': '#10B981',
  'Re-Engagement 90d': '#8B5CF6',
  'WhatsApp': '#25D366',
  'SMS Blast': '#06B6D4',
  'Personal Call': '#F97316',
  'Gift-with-Purchase': '#EC4899',
  'Surprise & Delight': '#14B8A6',
};

const LIFECYCLE_COLORS = {
  'New': '#3B82F6',
  'Active': '#10B981',
  'At-Risk': '#F59E0B',
  'Lapsed': '#EF4444',
};

const TIER_COLORS = {
  'Bronze': '#CD7F32',
  'Silver': '#C0C0C0',
  'Gold': '#FFD700',
  'Platinum': '#E5E4E2',
};

// ─── DEMO DATA ───
const DEMO_REVENUE = [
  { week: "2025-01-06", totalRevenue: 142500, subscriptionRevenue: 98200, oneTimeRevenue: 44300, refunds: 3200, netRevenue: 139300, totalOrders: 1850, aov: 77.03 },
  { week: "2025-01-13", totalRevenue: 148900, subscriptionRevenue: 103400, oneTimeRevenue: 45500, refunds: 2800, netRevenue: 146100, totalOrders: 1920, aov: 77.55 },
  { week: "2025-01-20", totalRevenue: 151200, subscriptionRevenue: 106800, oneTimeRevenue: 44400, refunds: 3500, netRevenue: 147700, totalOrders: 1960, aov: 77.14 },
  { week: "2025-01-27", totalRevenue: 155800, subscriptionRevenue: 110200, oneTimeRevenue: 45600, refunds: 2900, netRevenue: 152900, totalOrders: 2010, aov: 77.51 },
  { week: "2025-02-03", totalRevenue: 158400, subscriptionRevenue: 113500, oneTimeRevenue: 44900, refunds: 3100, netRevenue: 155300, totalOrders: 2040, aov: 77.65 },
  { week: "2025-02-10", totalRevenue: 162100, subscriptionRevenue: 116800, oneTimeRevenue: 45300, refunds: 2700, netRevenue: 159400, totalOrders: 2080, aov: 77.93 },
  { week: "2025-02-17", totalRevenue: 159800, subscriptionRevenue: 114200, oneTimeRevenue: 45600, refunds: 3400, netRevenue: 156400, totalOrders: 2050, aov: 77.95 },
  { week: "2025-02-24", totalRevenue: 165300, subscriptionRevenue: 119100, oneTimeRevenue: 46200, refunds: 2600, netRevenue: 162700, totalOrders: 2120, aov: 77.97 },
  { week: "2025-03-03", totalRevenue: 168900, subscriptionRevenue: 122400, oneTimeRevenue: 46500, refunds: 3000, netRevenue: 165900, totalOrders: 2170, aov: 77.83 },
  { week: "2025-03-10", totalRevenue: 172500, subscriptionRevenue: 125800, oneTimeRevenue: 46700, refunds: 2500, netRevenue: 170000, totalOrders: 2210, aov: 78.05 },
  { week: "2025-03-17", totalRevenue: 170100, subscriptionRevenue: 123600, oneTimeRevenue: 46500, refunds: 3300, netRevenue: 166800, totalOrders: 2180, aov: 78.03 },
  { week: "2025-03-24", totalRevenue: 175200, subscriptionRevenue: 128400, oneTimeRevenue: 46800, refunds: 2800, netRevenue: 172400, totalOrders: 2250, aov: 77.87 },
];

const DEMO_SUBSCRIPTIONS = [
  { month: "2024-10", activeSubscribers: 4120, newSubscribers: 580, churnedSubscribers: 310, reactivated: 45, mrr: 89200, churnRate: 7.5, voluntaryChurn: 5.2, involuntaryChurn: 2.3, ltv: 245, skipCount: 180 },
  { month: "2024-11", activeSubscribers: 4435, newSubscribers: 620, churnedSubscribers: 350, reactivated: 55, mrr: 96400, churnRate: 7.9, voluntaryChurn: 5.5, involuntaryChurn: 2.4, ltv: 252, skipCount: 195 },
  { month: "2024-12", activeSubscribers: 4780, newSubscribers: 710, churnedSubscribers: 420, reactivated: 65, mrr: 104800, churnRate: 8.8, voluntaryChurn: 6.1, involuntaryChurn: 2.7, ltv: 248, skipCount: 230 },
  { month: "2025-01", activeSubscribers: 5050, newSubscribers: 640, churnedSubscribers: 420, reactivated: 50, mrr: 110200, churnRate: 8.3, voluntaryChurn: 5.8, involuntaryChurn: 2.5, ltv: 258, skipCount: 210 },
  { month: "2025-02", activeSubscribers: 5280, newSubscribers: 590, churnedSubscribers: 400, reactivated: 40, mrr: 116800, churnRate: 7.6, voluntaryChurn: 5.3, involuntaryChurn: 2.3, ltv: 265, skipCount: 195 },
  { month: "2025-03", activeSubscribers: 5520, newSubscribers: 620, churnedSubscribers: 420, reactivated: 40, mrr: 125800, churnRate: 7.6, voluntaryChurn: 5.1, involuntaryChurn: 2.5, ltv: 272, skipCount: 205 },
];

const DEMO_EMAIL_FLOWS = [
  { week: "2025-01-06", type: "Campaign", flowName: null, sends: 42000, delivered: 40320, opens: 16530, openRate: 41.0, clicks: 2260, ctr: 5.61, unsubscribes: 84, unsubRate: 0.21, revenue: 7600, conversions: 195, listSize: 48500 },
  { week: "2025-01-06", type: "Flow", flowName: "Welcome Series", sends: 980, delivered: 960, opens: 576, openRate: 60.0, clicks: 211, ctr: 21.98, unsubscribes: 5, unsubRate: 0.52, revenue: 4200, conversions: 68 },
  { week: "2025-01-06", type: "Flow", flowName: "Abandoned Cart", sends: 2800, delivered: 2716, opens: 1358, openRate: 50.0, clicks: 435, ctr: 16.01, unsubscribes: 8, unsubRate: 0.29, revenue: 10200, conversions: 142 },
  { week: "2025-01-06", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1600, delivered: 1568, opens: 830, openRate: 52.9, clicks: 267, ctr: 17.03, unsubscribes: 3, unsubRate: 0.19, revenue: 1850, conversions: 45 },
  { week: "2025-01-06", type: "Flow", flowName: "Win-Back 60d", sends: 2400, delivered: 2304, opens: 691, openRate: 30.0, clicks: 138, ctr: 5.99, unsubscribes: 24, unsubRate: 1.04, revenue: 2800, conversions: 38 },
  { week: "2025-01-13", type: "Campaign", flowName: null, sends: 43000, delivered: 41280, opens: 17334, openRate: 42.0, clicks: 2395, ctr: 5.80, unsubscribes: 78, unsubRate: 0.19, revenue: 8100, conversions: 210, listSize: 49200 },
  { week: "2025-01-13", type: "Flow", flowName: "Welcome Series", sends: 1050, delivered: 1029, opens: 617, openRate: 59.9, clicks: 226, ctr: 21.96, unsubscribes: 4, unsubRate: 0.39, revenue: 4500, conversions: 72 },
  { week: "2025-01-13", type: "Flow", flowName: "Abandoned Cart", sends: 2900, delivered: 2813, opens: 1407, openRate: 50.0, clicks: 450, ctr: 16.00, unsubscribes: 9, unsubRate: 0.32, revenue: 10800, conversions: 148 },
  { week: "2025-01-13", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1650, delivered: 1617, opens: 873, openRate: 54.0, clicks: 283, ctr: 17.50, unsubscribes: 3, unsubRate: 0.19, revenue: 1950, conversions: 48 },
  { week: "2025-01-13", type: "Flow", flowName: "Win-Back 60d", sends: 2350, delivered: 2256, opens: 677, openRate: 30.0, clicks: 135, ctr: 5.98, unsubscribes: 22, unsubRate: 0.98, revenue: 2650, conversions: 36 },
  { week: "2025-01-20", type: "Campaign", flowName: null, sends: 43500, delivered: 41760, opens: 17546, openRate: 42.0, clicks: 2464, ctr: 5.90, unsubscribes: 72, unsubRate: 0.17, revenue: 8400, conversions: 218, listSize: 49800 },
  { week: "2025-01-20", type: "Flow", flowName: "Welcome Series", sends: 1100, delivered: 1078, opens: 647, openRate: 60.0, clicks: 237, ctr: 21.98, unsubscribes: 5, unsubRate: 0.46, revenue: 4800, conversions: 76 },
  { week: "2025-01-20", type: "Flow", flowName: "Abandoned Cart", sends: 3000, delivered: 2910, opens: 1455, openRate: 50.0, clicks: 466, ctr: 16.01, unsubscribes: 10, unsubRate: 0.34, revenue: 11200, conversions: 152 },
  { week: "2025-01-20", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1700, delivered: 1666, opens: 900, openRate: 54.0, clicks: 292, ctr: 17.53, unsubscribes: 3, unsubRate: 0.18, revenue: 2050, conversions: 50 },
  { week: "2025-01-20", type: "Flow", flowName: "Win-Back 60d", sends: 2300, delivered: 2208, opens: 663, openRate: 30.0, clicks: 132, ctr: 5.98, unsubscribes: 21, unsubRate: 0.95, revenue: 2500, conversions: 34 },
  { week: "2025-01-27", type: "Campaign", flowName: null, sends: 44000, delivered: 42240, opens: 17742, openRate: 42.0, clicks: 2535, ctr: 6.00, unsubscribes: 70, unsubRate: 0.17, revenue: 8800, conversions: 225, listSize: 50300 },
  { week: "2025-01-27", type: "Flow", flowName: "Welcome Series", sends: 1150, delivered: 1127, opens: 676, openRate: 60.0, clicks: 248, ctr: 22.0, unsubscribes: 5, unsubRate: 0.44, revenue: 5100, conversions: 80 },
  { week: "2025-01-27", type: "Flow", flowName: "Abandoned Cart", sends: 3100, delivered: 3007, opens: 1504, openRate: 50.0, clicks: 481, ctr: 15.99, unsubscribes: 11, unsubRate: 0.37, revenue: 11600, conversions: 158 },
  { week: "2025-01-27", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1750, delivered: 1715, opens: 926, openRate: 54.0, clicks: 300, ctr: 17.49, unsubscribes: 3, unsubRate: 0.17, revenue: 2150, conversions: 52 },
  { week: "2025-01-27", type: "Flow", flowName: "Win-Back 60d", sends: 2250, delivered: 2160, opens: 648, openRate: 30.0, clicks: 130, ctr: 6.02, unsubscribes: 20, unsubRate: 0.93, revenue: 2400, conversions: 32 },
  { week: "2025-02-03", type: "Campaign", flowName: null, sends: 44500, delivered: 42720, opens: 18342, openRate: 42.9, clicks: 2606, ctr: 6.10, unsubscribes: 68, unsubRate: 0.16, revenue: 9200, conversions: 235, listSize: 50800 },
  { week: "2025-02-03", type: "Flow", flowName: "Welcome Series", sends: 1200, delivered: 1176, opens: 706, openRate: 60.0, clicks: 259, ctr: 22.02, unsubscribes: 5, unsubRate: 0.43, revenue: 5400, conversions: 84 },
  { week: "2025-02-03", type: "Flow", flowName: "Abandoned Cart", sends: 3200, delivered: 3104, opens: 1552, openRate: 50.0, clicks: 497, ctr: 16.01, unsubscribes: 10, unsubRate: 0.32, revenue: 12000, conversions: 165 },
  { week: "2025-02-03", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1800, delivered: 1764, opens: 953, openRate: 54.0, clicks: 309, ctr: 17.52, unsubscribes: 3, unsubRate: 0.17, revenue: 2200, conversions: 54 },
  { week: "2025-02-03", type: "Flow", flowName: "Win-Back 60d", sends: 2200, delivered: 2112, opens: 634, openRate: 30.0, clicks: 127, ctr: 6.01, unsubscribes: 19, unsubRate: 0.90, revenue: 2300, conversions: 31 },
  { week: "2025-02-10", type: "Campaign", flowName: null, sends: 45000, delivered: 43200, opens: 18792, openRate: 43.5, clicks: 2678, ctr: 6.20, unsubscribes: 65, unsubRate: 0.15, revenue: 9500, conversions: 242, listSize: 51200 },
  { week: "2025-02-10", type: "Flow", flowName: "Welcome Series", sends: 1250, delivered: 1225, opens: 735, openRate: 60.0, clicks: 269, ctr: 21.96, unsubscribes: 5, unsubRate: 0.41, revenue: 5600, conversions: 88 },
  { week: "2025-02-10", type: "Flow", flowName: "Abandoned Cart", sends: 3300, delivered: 3201, opens: 1601, openRate: 50.0, clicks: 512, ctr: 15.99, unsubscribes: 10, unsubRate: 0.31, revenue: 12500, conversions: 170 },
  { week: "2025-02-10", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1850, delivered: 1813, opens: 979, openRate: 54.0, clicks: 317, ctr: 17.49, unsubscribes: 3, unsubRate: 0.17, revenue: 2300, conversions: 56 },
  { week: "2025-02-10", type: "Flow", flowName: "Win-Back 60d", sends: 2150, delivered: 2064, opens: 640, openRate: 31.0, clicks: 134, ctr: 6.49, unsubscribes: 18, unsubRate: 0.87, revenue: 2600, conversions: 35 },
  { week: "2025-02-17", type: "Campaign", flowName: null, sends: 45200, delivered: 43392, opens: 18636, openRate: 42.9, clicks: 2647, ctr: 6.10, unsubscribes: 67, unsubRate: 0.15, revenue: 9300, conversions: 238, listSize: 51500 },
  { week: "2025-02-17", type: "Flow", flowName: "Welcome Series", sends: 1180, delivered: 1156, opens: 694, openRate: 60.0, clicks: 254, ctr: 21.97, unsubscribes: 5, unsubRate: 0.43, revenue: 5300, conversions: 83 },
  { week: "2025-02-17", type: "Flow", flowName: "Abandoned Cart", sends: 3250, delivered: 3153, opens: 1576, openRate: 50.0, clicks: 504, ctr: 15.98, unsubscribes: 10, unsubRate: 0.32, revenue: 12200, conversions: 168 },
  { week: "2025-02-17", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1820, delivered: 1784, opens: 963, openRate: 54.0, clicks: 312, ctr: 17.49, unsubscribes: 3, unsubRate: 0.17, revenue: 2250, conversions: 55 },
  { week: "2025-02-17", type: "Flow", flowName: "Win-Back 60d", sends: 2100, delivered: 2016, opens: 645, openRate: 32.0, clicks: 141, ctr: 6.99, unsubscribes: 17, unsubRate: 0.84, revenue: 2800, conversions: 38 },
  { week: "2025-02-24", type: "Campaign", flowName: null, sends: 45500, delivered: 43680, opens: 19003, openRate: 43.5, clicks: 2737, ctr: 6.26, unsubscribes: 63, unsubRate: 0.14, revenue: 9800, conversions: 250, listSize: 51800 },
  { week: "2025-02-24", type: "Flow", flowName: "Welcome Series", sends: 1300, delivered: 1274, opens: 764, openRate: 60.0, clicks: 280, ctr: 21.98, unsubscribes: 5, unsubRate: 0.39, revenue: 5800, conversions: 91 },
  { week: "2025-02-24", type: "Flow", flowName: "Abandoned Cart", sends: 3350, delivered: 3250, opens: 1625, openRate: 50.0, clicks: 520, ctr: 16.0, unsubscribes: 11, unsubRate: 0.34, revenue: 13000, conversions: 175 },
  { week: "2025-02-24", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1880, delivered: 1842, opens: 995, openRate: 54.0, clicks: 322, ctr: 17.48, unsubscribes: 3, unsubRate: 0.16, revenue: 2350, conversions: 57 },
  { week: "2025-02-24", type: "Flow", flowName: "Win-Back 60d", sends: 2050, delivered: 1968, opens: 630, openRate: 32.0, clicks: 138, ctr: 7.01, unsubscribes: 16, unsubRate: 0.81, revenue: 2900, conversions: 40 },
  { week: "2025-03-03", type: "Campaign", flowName: null, sends: 46000, delivered: 44160, opens: 19218, openRate: 43.5, clicks: 2782, ctr: 6.30, unsubscribes: 60, unsubRate: 0.14, revenue: 10200, conversions: 260, listSize: 52200 },
  { week: "2025-03-03", type: "Flow", flowName: "Welcome Series", sends: 1350, delivered: 1323, opens: 794, openRate: 60.0, clicks: 291, ctr: 22.0, unsubscribes: 5, unsubRate: 0.38, revenue: 6000, conversions: 94 },
  { week: "2025-03-03", type: "Flow", flowName: "Abandoned Cart", sends: 3400, delivered: 3298, opens: 1649, openRate: 50.0, clicks: 528, ctr: 16.01, unsubscribes: 10, unsubRate: 0.30, revenue: 13400, conversions: 180 },
  { week: "2025-03-03", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1900, delivered: 1862, opens: 1006, openRate: 54.0, clicks: 326, ctr: 17.51, unsubscribes: 3, unsubRate: 0.16, revenue: 2400, conversions: 58 },
  { week: "2025-03-03", type: "Flow", flowName: "Win-Back 60d", sends: 2000, delivered: 1920, opens: 634, openRate: 33.0, clicks: 144, ctr: 7.50, unsubscribes: 15, unsubRate: 0.78, revenue: 3100, conversions: 42 },
  { week: "2025-03-03", type: "Flow", flowName: "Re-Engagement 90d", sends: 1800, delivered: 1728, opens: 449, openRate: 26.0, clicks: 69, ctr: 3.99, unsubscribes: 36, unsubRate: 2.08, revenue: 800, conversions: 12 },
  { week: "2025-03-10", type: "Campaign", flowName: null, sends: 46500, delivered: 44640, opens: 19436, openRate: 43.5, clicks: 2812, ctr: 6.30, unsubscribes: 58, unsubRate: 0.13, revenue: 10500, conversions: 268, listSize: 52600 },
  { week: "2025-03-10", type: "Flow", flowName: "Welcome Series", sends: 1400, delivered: 1372, opens: 823, openRate: 60.0, clicks: 302, ctr: 22.01, unsubscribes: 5, unsubRate: 0.36, revenue: 6200, conversions: 97 },
  { week: "2025-03-10", type: "Flow", flowName: "Abandoned Cart", sends: 3500, delivered: 3395, opens: 1698, openRate: 50.0, clicks: 543, ctr: 15.99, unsubscribes: 10, unsubRate: 0.29, revenue: 13800, conversions: 186 },
  { week: "2025-03-10", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1950, delivered: 1911, opens: 1032, openRate: 54.0, clicks: 334, ctr: 17.48, unsubscribes: 3, unsubRate: 0.16, revenue: 2500, conversions: 60 },
  { week: "2025-03-10", type: "Flow", flowName: "Win-Back 60d", sends: 1950, delivered: 1872, opens: 618, openRate: 33.0, clicks: 140, ctr: 7.48, unsubscribes: 14, unsubRate: 0.75, revenue: 3200, conversions: 44 },
  { week: "2025-03-10", type: "Flow", flowName: "Re-Engagement 90d", sends: 1750, delivered: 1680, opens: 470, openRate: 28.0, clicks: 84, ctr: 5.0, unsubscribes: 32, unsubRate: 1.90, revenue: 1100, conversions: 16 },
  { week: "2025-03-17", type: "Campaign", flowName: null, sends: 46800, delivered: 44928, opens: 19571, openRate: 43.6, clicks: 2830, ctr: 6.30, unsubscribes: 56, unsubRate: 0.12, revenue: 10800, conversions: 275, listSize: 53000 },
  { week: "2025-03-17", type: "Flow", flowName: "Welcome Series", sends: 1420, delivered: 1392, opens: 835, openRate: 60.0, clicks: 306, ctr: 21.98, unsubscribes: 5, unsubRate: 0.36, revenue: 6400, conversions: 100 },
  { week: "2025-03-17", type: "Flow", flowName: "Abandoned Cart", sends: 3550, delivered: 3444, opens: 1722, openRate: 50.0, clicks: 551, ctr: 16.0, unsubscribes: 10, unsubRate: 0.29, revenue: 14000, conversions: 190 },
  { week: "2025-03-17", type: "Flow", flowName: "Post-Purchase Upsell", sends: 1980, delivered: 1940, opens: 1048, openRate: 54.0, clicks: 339, ctr: 17.47, unsubscribes: 3, unsubRate: 0.15, revenue: 2550, conversions: 62 },
  { week: "2025-03-17", type: "Flow", flowName: "Win-Back 60d", sends: 1900, delivered: 1824, opens: 602, openRate: 33.0, clicks: 137, ctr: 7.51, unsubscribes: 13, unsubRate: 0.71, revenue: 3300, conversions: 45 },
  { week: "2025-03-17", type: "Flow", flowName: "Re-Engagement 90d", sends: 1700, delivered: 1632, opens: 490, openRate: 30.0, clicks: 98, ctr: 6.0, unsubscribes: 28, unsubRate: 1.72, revenue: 1400, conversions: 20 },
  { week: "2025-03-24", type: "Campaign", flowName: null, sends: 47000, delivered: 45120, opens: 19853, openRate: 44.0, clicks: 2888, ctr: 6.40, unsubscribes: 52, unsubRate: 0.12, revenue: 11200, conversions: 285, listSize: 53400 },
  { week: "2025-03-24", type: "Flow", flowName: "Welcome Series", sends: 1450, delivered: 1421, opens: 853, openRate: 60.0, clicks: 313, ctr: 22.03, unsubscribes: 5, unsubRate: 0.35, revenue: 6600, conversions: 103 },
  { week: "2025-03-24", type: "Flow", flowName: "Abandoned Cart", sends: 3600, delivered: 3492, opens: 1746, openRate: 50.0, clicks: 559, ctr: 16.01, unsubscribes: 10, unsubRate: 0.29, revenue: 14200, conversions: 192 },
  { week: "2025-03-24", type: "Flow", flowName: "Post-Purchase Upsell", sends: 2000, delivered: 1960, opens: 1058, openRate: 54.0, clicks: 343, ctr: 17.50, unsubscribes: 3, unsubRate: 0.15, revenue: 2600, conversions: 63 },
  { week: "2025-03-24", type: "Flow", flowName: "Win-Back 60d", sends: 1850, delivered: 1776, opens: 604, openRate: 34.0, clicks: 142, ctr: 7.99, unsubscribes: 12, unsubRate: 0.68, revenue: 3500, conversions: 48 },
  { week: "2025-03-24", type: "Flow", flowName: "Re-Engagement 90d", sends: 1650, delivered: 1584, opens: 507, openRate: 32.0, clicks: 111, ctr: 7.01, unsubscribes: 25, unsubRate: 1.58, revenue: 1700, conversions: 24 },
];

const DEMO_LOYALTY = [
  { month: "2024-10", totalMembers: 820, newEnrollments: 820, pointsIssued: 164000, pointsRedeemed: 8200, redemptionRate: 5.0, rewardsRedeemed: 41, rewardCostGBP: 1230, revenueFromMembers: 28500, revenueFromNonMembers: 114000, memberAOV: 82.50, nonMemberAOV: 74.20, memberRetentionRate: 88.0, nonMemberRetentionRate: 78.0, tierBronze: 720, tierSilver: 80, tierGold: 18, tierPlatinum: 2 },
  { month: "2024-11", totalMembers: 1380, newEnrollments: 560, pointsIssued: 276000, pointsRedeemed: 27600, redemptionRate: 10.0, rewardsRedeemed: 92, rewardCostGBP: 2760, revenueFromMembers: 52200, revenueFromNonMembers: 100200, memberAOV: 84.00, nonMemberAOV: 73.80, memberRetentionRate: 89.5, nonMemberRetentionRate: 76.0, tierBronze: 1150, tierSilver: 175, tierGold: 45, tierPlatinum: 10 },
  { month: "2024-12", totalMembers: 1950, newEnrollments: 570, pointsIssued: 390000, pointsRedeemed: 54600, redemptionRate: 14.0, rewardsRedeemed: 182, rewardCostGBP: 5460, revenueFromMembers: 78400, revenueFromNonMembers: 82400, memberAOV: 85.50, nonMemberAOV: 73.50, memberRetentionRate: 90.2, nonMemberRetentionRate: 72.0, tierBronze: 1520, tierSilver: 310, tierGold: 95, tierPlatinum: 25 },
  { month: "2025-01", totalMembers: 2540, newEnrollments: 590, pointsIssued: 508000, pointsRedeemed: 81280, redemptionRate: 16.0, rewardsRedeemed: 254, rewardCostGBP: 7620, revenueFromMembers: 105600, revenueFromNonMembers: 72400, memberAOV: 86.80, nonMemberAOV: 73.20, memberRetentionRate: 91.0, nonMemberRetentionRate: 71.0, tierBronze: 1880, tierSilver: 440, tierGold: 170, tierPlatinum: 50 },
  { month: "2025-02", totalMembers: 3100, newEnrollments: 560, pointsIssued: 620000, pointsRedeemed: 111600, redemptionRate: 18.0, rewardsRedeemed: 341, rewardCostGBP: 10230, revenueFromMembers: 134200, revenueFromNonMembers: 58800, memberAOV: 88.20, nonMemberAOV: 72.80, memberRetentionRate: 91.8, nonMemberRetentionRate: 70.5, tierBronze: 2200, tierSilver: 580, tierGold: 245, tierPlatinum: 75 },
  { month: "2025-03", totalMembers: 3680, newEnrollments: 580, pointsIssued: 736000, pointsRedeemed: 147200, redemptionRate: 20.0, rewardsRedeemed: 442, rewardCostGBP: 13260, revenueFromMembers: 162400, revenueFromNonMembers: 48200, memberAOV: 89.50, nonMemberAOV: 72.50, memberRetentionRate: 92.4, nonMemberRetentionRate: 69.8, tierBronze: 2520, tierSilver: 720, tierGold: 330, tierPlatinum: 110 },
];

const DEMO_SEGMENTS = [
  { month: "2024-10", segNew: 580, segActive: 3200, segAtRisk: 1450, segLapsed: 2800, totalCustomers: 8030, avgRFMScore: 2.8, segNewRevenue: 25200, segActiveRevenue: 92800, segAtRiskRevenue: 14500, segLapsedRevenue: 2800, migratedAtRiskToActive: 85, migratedActiveToAtRisk: 145, reactivatedFromLapsed: 45, avgOrdersPerActiveCustomer: 2.1 },
  { month: "2024-11", segNew: 620, segActive: 3350, segAtRisk: 1380, segLapsed: 2780, totalCustomers: 8130, avgRFMScore: 2.9, segNewRevenue: 27900, segActiveRevenue: 98200, segAtRiskRevenue: 13800, segLapsedRevenue: 2780, migratedAtRiskToActive: 110, migratedActiveToAtRisk: 120, reactivatedFromLapsed: 55, avgOrdersPerActiveCustomer: 2.2 },
  { month: "2024-12", segNew: 710, segActive: 3520, segAtRisk: 1310, segLapsed: 2720, totalCustomers: 8260, avgRFMScore: 3.0, segNewRevenue: 33500, segActiveRevenue: 105600, segAtRiskRevenue: 13100, segLapsedRevenue: 2720, migratedAtRiskToActive: 130, migratedActiveToAtRisk: 100, reactivatedFromLapsed: 65, avgOrdersPerActiveCustomer: 2.2 },
  { month: "2025-01", segNew: 640, segActive: 3750, segAtRisk: 1240, segLapsed: 2650, totalCustomers: 8280, avgRFMScore: 3.1, segNewRevenue: 29400, segActiveRevenue: 112500, segAtRiskRevenue: 12400, segLapsedRevenue: 2650, migratedAtRiskToActive: 145, migratedActiveToAtRisk: 95, reactivatedFromLapsed: 50, avgOrdersPerActiveCustomer: 2.3 },
  { month: "2025-02", segNew: 590, segActive: 3920, segAtRisk: 1160, segLapsed: 2580, totalCustomers: 8250, avgRFMScore: 3.2, segNewRevenue: 27600, segActiveRevenue: 121500, segAtRiskRevenue: 11600, segLapsedRevenue: 2580, migratedAtRiskToActive: 160, migratedActiveToAtRisk: 80, reactivatedFromLapsed: 40, avgOrdersPerActiveCustomer: 2.3 },
  { month: "2025-03", segNew: 620, segActive: 4100, segAtRisk: 1080, segLapsed: 2500, totalCustomers: 8300, avgRFMScore: 3.3, segNewRevenue: 29800, segActiveRevenue: 131200, segAtRiskRevenue: 10800, segLapsedRevenue: 2500, migratedAtRiskToActive: 175, migratedActiveToAtRisk: 70, reactivatedFromLapsed: 40, avgOrdersPerActiveCustomer: 2.4 },
];

const DEMO_OUTREACH = [
  { week: "2025-01-06", channel: "WhatsApp", sends: 3200, delivered: 3104, responses: 1242, responseRate: 40.0, conversions: 128, conversionRate: 4.12, revenue: 9800, cost: 320 },
  { week: "2025-01-06", channel: "SMS Blast", sends: 8500, delivered: 8245, responses: 742, responseRate: 9.0, conversions: 195, conversionRate: 2.37, revenue: 11200, cost: 680 },
  { week: "2025-01-06", channel: "Personal Call", sends: 45, delivered: 45, responses: 32, responseRate: 71.1, conversions: 18, conversionRate: 40.0, revenue: 4200, cost: 540 },
  { week: "2025-01-13", channel: "WhatsApp", sends: 3400, delivered: 3298, responses: 1319, responseRate: 40.0, conversions: 136, conversionRate: 4.12, revenue: 10400, cost: 340 },
  { week: "2025-01-13", channel: "SMS Blast", sends: 8800, delivered: 8536, responses: 768, responseRate: 9.0, conversions: 202, conversionRate: 2.37, revenue: 11600, cost: 704 },
  { week: "2025-01-13", channel: "Personal Call", sends: 48, delivered: 48, responses: 34, responseRate: 70.8, conversions: 19, conversionRate: 39.6, revenue: 4500, cost: 576 },
  { week: "2025-01-13", channel: "Gift-with-Purchase", sends: 250, delivered: 250, responses: 250, responseRate: 100.0, conversions: 250, conversionRate: 100.0, revenue: 6200, cost: 2500 },
  { week: "2025-01-20", channel: "WhatsApp", sends: 3500, delivered: 3395, responses: 1358, responseRate: 40.0, conversions: 142, conversionRate: 4.18, revenue: 10800, cost: 350 },
  { week: "2025-01-20", channel: "SMS Blast", sends: 9000, delivered: 8730, responses: 786, responseRate: 9.0, conversions: 208, conversionRate: 2.38, revenue: 12000, cost: 720 },
  { week: "2025-01-20", channel: "Personal Call", sends: 50, delivered: 50, responses: 36, responseRate: 72.0, conversions: 20, conversionRate: 40.0, revenue: 4800, cost: 600 },
  { week: "2025-01-27", channel: "WhatsApp", sends: 3600, delivered: 3492, responses: 1397, responseRate: 40.0, conversions: 148, conversionRate: 4.24, revenue: 11200, cost: 360 },
  { week: "2025-01-27", channel: "SMS Blast", sends: 9200, delivered: 8924, responses: 803, responseRate: 9.0, conversions: 215, conversionRate: 2.41, revenue: 12400, cost: 736 },
  { week: "2025-01-27", channel: "Personal Call", sends: 52, delivered: 52, responses: 37, responseRate: 71.2, conversions: 21, conversionRate: 40.4, revenue: 5100, cost: 624 },
  { week: "2025-01-27", channel: "Surprise & Delight", sends: 100, delivered: 100, responses: 85, responseRate: 85.0, conversions: 72, conversionRate: 72.0, revenue: 3600, cost: 1500 },
  { week: "2025-02-03", channel: "WhatsApp", sends: 3800, delivered: 3686, responses: 1474, responseRate: 40.0, conversions: 156, conversionRate: 4.23, revenue: 11800, cost: 380 },
  { week: "2025-02-03", channel: "SMS Blast", sends: 12000, delivered: 11640, responses: 1280, responseRate: 11.0, conversions: 310, conversionRate: 2.66, revenue: 18500, cost: 960 },
  { week: "2025-02-03", channel: "Personal Call", sends: 55, delivered: 55, responses: 39, responseRate: 70.9, conversions: 22, conversionRate: 40.0, revenue: 5400, cost: 660 },
  { week: "2025-02-10", channel: "WhatsApp", sends: 3900, delivered: 3783, responses: 1551, responseRate: 41.0, conversions: 162, conversionRate: 4.28, revenue: 12200, cost: 390 },
  { week: "2025-02-10", channel: "SMS Blast", sends: 11500, delivered: 11155, responses: 1227, responseRate: 11.0, conversions: 298, conversionRate: 2.67, revenue: 17800, cost: 920 },
  { week: "2025-02-10", channel: "Personal Call", sends: 55, delivered: 55, responses: 40, responseRate: 72.7, conversions: 23, conversionRate: 41.8, revenue: 5600, cost: 660 },
  { week: "2025-02-10", channel: "Gift-with-Purchase", sends: 280, delivered: 280, responses: 280, responseRate: 100.0, conversions: 280, conversionRate: 100.0, revenue: 7000, cost: 2800 },
  { week: "2025-02-17", channel: "WhatsApp", sends: 4000, delivered: 3880, responses: 1590, responseRate: 41.0, conversions: 168, conversionRate: 4.33, revenue: 12600, cost: 400 },
  { week: "2025-02-17", channel: "SMS Blast", sends: 11200, delivered: 10864, responses: 1195, responseRate: 11.0, conversions: 290, conversionRate: 2.67, revenue: 17200, cost: 896 },
  { week: "2025-02-17", channel: "Personal Call", sends: 58, delivered: 58, responses: 42, responseRate: 72.4, conversions: 24, conversionRate: 41.4, revenue: 5800, cost: 696 },
  { week: "2025-02-24", channel: "WhatsApp", sends: 4100, delivered: 3977, responses: 1631, responseRate: 41.0, conversions: 174, conversionRate: 4.38, revenue: 13000, cost: 410 },
  { week: "2025-02-24", channel: "SMS Blast", sends: 11000, delivered: 10670, responses: 1174, responseRate: 11.0, conversions: 285, conversionRate: 2.67, revenue: 16800, cost: 880 },
  { week: "2025-02-24", channel: "Personal Call", sends: 58, delivered: 58, responses: 42, responseRate: 72.4, conversions: 24, conversionRate: 41.4, revenue: 5900, cost: 696 },
  { week: "2025-02-24", channel: "Surprise & Delight", sends: 120, delivered: 120, responses: 102, responseRate: 85.0, conversions: 86, conversionRate: 71.7, revenue: 4300, cost: 1800 },
  { week: "2025-03-03", channel: "WhatsApp", sends: 4200, delivered: 4074, responses: 1711, responseRate: 42.0, conversions: 182, conversionRate: 4.47, revenue: 13600, cost: 420 },
  { week: "2025-03-03", channel: "SMS Blast", sends: 10800, delivered: 10476, responses: 1152, responseRate: 11.0, conversions: 280, conversionRate: 2.67, revenue: 16500, cost: 864 },
  { week: "2025-03-03", channel: "Personal Call", sends: 60, delivered: 60, responses: 43, responseRate: 71.7, conversions: 25, conversionRate: 41.7, revenue: 6100, cost: 720 },
  { week: "2025-03-10", channel: "WhatsApp", sends: 4300, delivered: 4171, responses: 1752, responseRate: 42.0, conversions: 188, conversionRate: 4.51, revenue: 14000, cost: 430 },
  { week: "2025-03-10", channel: "SMS Blast", sends: 10500, delivered: 10185, responses: 1120, responseRate: 11.0, conversions: 272, conversionRate: 2.67, revenue: 16000, cost: 840 },
  { week: "2025-03-10", channel: "Personal Call", sends: 60, delivered: 60, responses: 44, responseRate: 73.3, conversions: 26, conversionRate: 43.3, revenue: 6300, cost: 720 },
  { week: "2025-03-10", channel: "Gift-with-Purchase", sends: 300, delivered: 300, responses: 300, responseRate: 100.0, conversions: 300, conversionRate: 100.0, revenue: 7500, cost: 3000 },
  { week: "2025-03-17", channel: "WhatsApp", sends: 4400, delivered: 4268, responses: 1793, responseRate: 42.0, conversions: 194, conversionRate: 4.55, revenue: 14400, cost: 440 },
  { week: "2025-03-17", channel: "SMS Blast", sends: 10200, delivered: 9894, responses: 1088, responseRate: 11.0, conversions: 265, conversionRate: 2.68, revenue: 15600, cost: 816 },
  { week: "2025-03-17", channel: "Personal Call", sends: 62, delivered: 62, responses: 45, responseRate: 72.6, conversions: 26, conversionRate: 41.9, revenue: 6400, cost: 744 },
  { week: "2025-03-24", channel: "WhatsApp", sends: 4500, delivered: 4365, responses: 1833, responseRate: 42.0, conversions: 200, conversionRate: 4.58, revenue: 14800, cost: 450 },
  { week: "2025-03-24", channel: "SMS Blast", sends: 10000, delivered: 9700, responses: 1067, responseRate: 11.0, conversions: 260, conversionRate: 2.68, revenue: 15300, cost: 800 },
  { week: "2025-03-24", channel: "Personal Call", sends: 65, delivered: 65, responses: 47, responseRate: 72.3, conversions: 28, conversionRate: 43.1, revenue: 6800, cost: 780 },
  { week: "2025-03-24", channel: "Surprise & Delight", sends: 130, delivered: 130, responses: 110, responseRate: 84.6, conversions: 93, conversionRate: 71.5, revenue: 4650, cost: 1950 },
];

const DEMO_BEFORE_AFTER = [
  { activity: "Loyalty Program Launch", launchDate: "2024-10-01", metric: "Monthly Retention Rate", beforeValue: 78.0, afterValue: 92.4, beforePeriod: "Jul-Sep 2024", afterPeriod: "Mar 2025", lift: 18.5, unit: "percent" },
  { activity: "Loyalty Program Launch", launchDate: "2024-10-01", metric: "Average Order Value", beforeValue: 74.20, afterValue: 89.50, beforePeriod: "Jul-Sep 2024", afterPeriod: "Mar 2025", lift: 20.6, unit: "currency" },
  { activity: "Win-Back SMS Campaign", launchDate: "2025-02-03", metric: "Lapsed Reactivation Rate", beforeValue: 1.6, afterValue: 3.8, beforePeriod: "Dec 2024-Jan 2025", afterPeriod: "Feb-Mar 2025", lift: 137.5, unit: "percent" },
  { activity: "Win-Back SMS Campaign", launchDate: "2025-02-03", metric: "Win-Back Revenue per Week", beforeValue: 2600, afterValue: 3500, beforePeriod: "Dec 2024-Jan 2025", afterPeriod: "Feb-Mar 2025", lift: 34.6, unit: "currency" },
  { activity: "Re-Engagement Flow Launch", launchDate: "2025-03-03", metric: "At-Risk Save Rate", beforeValue: 5.9, afterValue: 16.2, beforePeriod: "Jan-Feb 2025", afterPeriod: "Mar 2025", lift: 174.6, unit: "percent" },
  { activity: "WhatsApp Channel Expansion", launchDate: "2025-01-06", metric: "Direct Outreach Revenue/Week", beforeValue: 8200, afterValue: 26250, beforePeriod: "Nov-Dec 2024", afterPeriod: "Mar 2025 avg", lift: 220.1, unit: "currency" },
];

const DEMO_HOLDOUT_TESTS = [
  { testName: "Abandoned Cart Flow", testPeriod: "Jan-Mar 2025", sampleSize: 42000, controlSize: 4200, exposedSize: 37800, controlConversionRate: 8.2, exposedConversionRate: 16.0, controlRevPerCustomer: 18.50, exposedRevPerCustomer: 38.20, incrementalConversionLift: 95.1, incrementalRevLift: 106.5, incrementalRevenue: 48200, confidence: 0.96, status: "active" },
  { testName: "Welcome Series Flow", testPeriod: "Jan-Mar 2025", sampleSize: 14000, controlSize: 1400, exposedSize: 12600, controlConversionRate: 42.0, exposedConversionRate: 60.0, controlRevPerCustomer: 52.00, exposedRevPerCustomer: 78.40, incrementalConversionLift: 42.9, incrementalRevLift: 50.8, incrementalRevenue: 32400, confidence: 0.94, status: "active" },
  { testName: "Loyalty Points Emails", testPeriod: "Jan-Mar 2025", sampleSize: 18000, controlSize: 1800, exposedSize: 16200, controlConversionRate: 12.5, exposedConversionRate: 18.8, controlRevPerCustomer: 24.00, exposedRevPerCustomer: 36.80, incrementalConversionLift: 50.4, incrementalRevLift: 53.3, incrementalRevenue: 22100, confidence: 0.91, status: "active" },
  { testName: "WhatsApp Win-Back Nudge", testPeriod: "Feb-Mar 2025", sampleSize: 6000, controlSize: 1500, exposedSize: 4500, controlConversionRate: 2.1, exposedConversionRate: 4.5, controlRevPerCustomer: 4.20, exposedRevPerCustomer: 12.80, incrementalConversionLift: 114.3, incrementalRevLift: 204.8, incrementalRevenue: 8600, confidence: 0.87, status: "active" },
  { testName: "Post-Purchase Upsell Flow", testPeriod: "Jan-Mar 2025", sampleSize: 22000, controlSize: 2200, exposedSize: 19800, controlConversionRate: 5.8, exposedConversionRate: 17.5, controlRevPerCustomer: 8.20, exposedRevPerCustomer: 28.40, incrementalConversionLift: 201.7, incrementalRevLift: 246.3, incrementalRevenue: 28500, confidence: 0.95, status: "active" },
];

const DEMO_ACTIVITY_ROI = [
  { activity: "Welcome Series Flow", channel: "Email", totalCost: 480, attributedRevenue: 66000, incrementalRevenue: 32400, incrementalROI: 66.5, customersInfluenced: 12600, period: "Q1 2025" },
  { activity: "Abandoned Cart Flow", channel: "Email", totalCost: 620, attributedRevenue: 152000, incrementalRevenue: 48200, incrementalROI: 76.7, customersInfluenced: 37800, period: "Q1 2025" },
  { activity: "Post-Purchase Upsell Flow", channel: "Email", totalCost: 380, attributedRevenue: 28200, incrementalRevenue: 28500, incrementalROI: 74.0, customersInfluenced: 19800, period: "Q1 2025" },
  { activity: "Win-Back 60d Flow", channel: "Email", totalCost: 350, attributedRevenue: 35400, incrementalRevenue: 14200, incrementalROI: 39.6, customersInfluenced: 6200, period: "Q1 2025" },
  { activity: "Re-Engagement 90d Flow", channel: "Email", totalCost: 180, attributedRevenue: 5000, incrementalRevenue: 3200, incrementalROI: 16.8, customersInfluenced: 5064, period: "Mar 2025" },
  { activity: "Loyalty Program", channel: "Cross-Channel", totalCost: 13260, attributedRevenue: 162400, incrementalRevenue: 42000, incrementalROI: 2.2, customersInfluenced: 3680, period: "Q1 2025" },
  { activity: "WhatsApp Campaigns", channel: "WhatsApp", totalCost: 4440, attributedRevenue: 148200, incrementalRevenue: 62800, incrementalROI: 13.1, customersInfluenced: 48600, period: "Q1 2025" },
  { activity: "SMS Blasts", channel: "SMS", totalCost: 9016, attributedRevenue: 179400, incrementalRevenue: 54200, incrementalROI: 5.0, customersInfluenced: 126400, period: "Q1 2025" },
  { activity: "Personal Calls", channel: "Phone", totalCost: 8016, attributedRevenue: 66900, incrementalRevenue: 45000, incrementalROI: 4.6, customersInfluenced: 715, period: "Q1 2025" },
  { activity: "Gift-with-Purchase", channel: "In-Box", totalCost: 8300, attributedRevenue: 20700, incrementalRevenue: 12400, incrementalROI: 0.5, customersInfluenced: 830, period: "Q1 2025" },
  { activity: "Surprise & Delight", channel: "In-Box", totalCost: 5250, attributedRevenue: 12550, incrementalRevenue: 8200, incrementalROI: 0.6, customersInfluenced: 350, period: "Q1 2025" },
];

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
const lastN = (arr, n) => arr?.slice(-n) ?? [];

// ─── TABS ───
const TABS = [
  { key: 'overview', label: 'CRM Overview' },
  { key: 'email', label: 'Email & Flows' },
  { key: 'loyalty', label: 'Loyalty & Rewards' },
  { key: 'segments', label: 'Segments & Lifecycle' },
  { key: 'outreach', label: 'Direct Outreach' },
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
  lastUpdated: { emailFlows: 'Demo', loyalty: 'Demo', segments: 'Demo', outreach: 'Demo', beforeAfter: 'Demo', holdoutTests: 'Demo', activityROI: 'Demo', revenue: 'Demo', subscriptions: 'Demo' },
  settingsOpen: false,
  tabPeriods: { overview: 'weekly', email: 'weekly', loyalty: 'monthly', segments: 'monthly', outreach: 'weekly', incrementality: 'all' },
};

function reducer(state, action) {
  switch (action.type) {
    case 'SET_TAB': return { ...state, activeTab: action.payload };
    case 'LOAD_DATA': return { ...state, [action.source]: action.payload, lastUpdated: { ...state.lastUpdated, [action.source]: new Date().toLocaleString('en-GB') } };
    case 'RESET_DEMO': return { ...initialState, activeTab: state.activeTab, settingsOpen: state.settingsOpen, tabPeriods: state.tabPeriods };
    case 'CLEAR_ALL': return { ...state, emailFlows: [], loyalty: [], segments: [], outreach: [], beforeAfter: [], holdoutTests: [], activityROI: [], revenue: [], subscriptions: [], lastUpdated: Object.fromEntries(Object.keys(state.lastUpdated).map(k => [k, 'Cleared'])) };
    case 'TOGGLE_SETTINGS': return { ...state, settingsOpen: !state.settingsOpen };
    case 'SET_TAB_PERIOD': return { ...state, tabPeriods: { ...state.tabPeriods, [action.tab]: action.period } };
    default: return state;
  }
}

// ─── CSV TEMPLATES ───
const CSV_TEMPLATES = {
  emailFlows: { headers: ['week','type','flowName','sends','delivered','opens','openRate','clicks','ctr','unsubscribes','unsubRate','revenue','conversions','listSize'], sample: [['2025-03-24','Campaign','','47000','45120','19853','44.0','2888','6.40','52','0.12','11200','285','53400']] },
  loyalty: { headers: ['month','totalMembers','newEnrollments','pointsIssued','pointsRedeemed','redemptionRate','rewardsRedeemed','rewardCostGBP','revenueFromMembers','revenueFromNonMembers','memberAOV','nonMemberAOV','memberRetentionRate','nonMemberRetentionRate','tierBronze','tierSilver','tierGold','tierPlatinum'], sample: [['2025-03','3680','580','736000','147200','20.0','442','13260','162400','48200','89.50','72.50','92.4','69.8','2520','720','330','110']] },
  segments: { headers: ['month','segNew','segActive','segAtRisk','segLapsed','totalCustomers','avgRFMScore','segNewRevenue','segActiveRevenue','segAtRiskRevenue','segLapsedRevenue','migratedAtRiskToActive','migratedActiveToAtRisk','reactivatedFromLapsed','avgOrdersPerActiveCustomer'], sample: [['2025-03','620','4100','1080','2500','8300','3.3','29800','131200','10800','2500','175','70','40','2.4']] },
  outreach: { headers: ['week','channel','sends','delivered','responses','responseRate','conversions','conversionRate','revenue','cost'], sample: [['2025-03-24','WhatsApp','4500','4365','1833','42.0','200','4.58','14800','450']] },
  beforeAfter: { headers: ['activity','launchDate','metric','beforeValue','afterValue','beforePeriod','afterPeriod','lift','unit'], sample: [['Loyalty Program Launch','2024-10-01','Monthly Retention Rate','78.0','92.4','Jul-Sep 2024','Mar 2025','18.5','percent']] },
  holdoutTests: { headers: ['testName','testPeriod','sampleSize','controlSize','exposedSize','controlConversionRate','exposedConversionRate','controlRevPerCustomer','exposedRevPerCustomer','incrementalConversionLift','incrementalRevLift','incrementalRevenue','confidence','status'], sample: [['Abandoned Cart Flow','Jan-Mar 2025','42000','4200','37800','8.2','16.0','18.50','38.20','95.1','106.5','48200','0.96','active']] },
  activityROI: { headers: ['activity','channel','totalCost','attributedRevenue','incrementalRevenue','incrementalROI','customersInfluenced','period'], sample: [['Welcome Series Flow','Email','480','66000','32400','66.5','12600','Q1 2025']] },
  revenue: { headers: ['week','totalRevenue','subscriptionRevenue','oneTimeRevenue','refunds','netRevenue','totalOrders','aov'], sample: [['2025-03-24','175200','128400','46800','2800','172400','2250','77.87']] },
  subscriptions: { headers: ['month','activeSubscribers','newSubscribers','churnedSubscribers','reactivated','mrr','churnRate','voluntaryChurn','involuntaryChurn','ltv','skipCount'], sample: [['2025-03','5520','620','420','40','125800','7.6','5.1','2.5','272','205']] },
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
    if (latest.redemptionRate < 10) alerts.push({ severity: 'warning', metric: 'Loyalty Redemption', value: latest.redemptionRate + '%', message: 'Below 10% — members not engaging' });
    const cumIssued = loy.reduce((s, m) => s + m.pointsIssued, 0);
    const cumRedeemed = loy.reduce((s, m) => s + m.pointsRedeemed, 0);
    const liability = (cumIssued - cumRedeemed) / 100;
    if (liability > 20000) alerts.push({ severity: 'warning', metric: 'Points Liability', value: formatCurrency(liability), message: 'Consider expiry policy' });
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
    <div style={{ background: C.cardBg, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: '10px 14px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
      <p style={{ margin: 0, fontWeight: 600, color: C.textPrimary, fontSize: 12 }}>{label}</p>
      {payload.map((p, i) => (
        <p key={i} style={{ margin: '4px 0 0', color: p.color || C.textSecondary, fontSize: 12 }}>
          {p.name}: {formatter ? formatter(p.value, p.name) : (typeof p.value === 'number' ? p.value.toLocaleString('en-GB') : p.value)}
        </p>
      ))}
    </div>
  );
}

function KPICard({ label, value, format = 'number', delta, status, sparkData, sparkKey, presentationMode }) {
  const fmt = (v) => {
    if (format === 'currency') return formatCurrency(v);
    if (format === 'currencyDecimal') return formatCurrencyDecimal(v);
    if (format === 'percent') return formatPercent(v);
    if (format === 'multiplier') return formatMultiplier(v);
    if (format === 'text') return v;
    return formatNumber(v);
  };
  const statusColor = status === 'good' ? C.success : status === 'warning' ? C.warning : status === 'danger' ? C.danger : C.textTertiary;
  return (
    <div style={{ background: C.cardBg, borderRadius: 12, padding: presentationMode ? '20px' : '16px', border: `1px solid ${C.cardBorder}`, display: 'flex', flexDirection: 'column', gap: 4 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span style={{ fontSize: presentationMode ? 13 : 11, color: C.textSecondary, fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.05em' }}>{label}</span>
        <span style={{ width: 8, height: 8, borderRadius: '50%', background: statusColor }} />
      </div>
      <span style={{ fontSize: presentationMode ? 28 : 22, fontWeight: 700, color: C.textPrimary, lineHeight: 1.1 }}>{fmt(value)}</span>
      {delta != null && (
        <span style={{ fontSize: 12, color: delta >= 0 ? C.success : C.danger, fontWeight: 600 }}>
          {delta >= 0 ? '▲' : '▼'} {Math.abs(delta).toFixed(1)}%
        </span>
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
    <div style={{ display: 'flex', gap: 4, overflowX: 'auto', padding: '8px 0', position: 'sticky', top: 0, zIndex: 10, background: C.pageBg, borderBottom: `1px solid ${C.cardBorder}` }}>
      {tabs.map(t => (
        <button key={t.key} onClick={() => onSelect(t.key)} style={{ padding: '8px 16px', borderRadius: 8, border: 'none', cursor: 'pointer', fontSize: 13, fontWeight: active === t.key ? 700 : 500, background: active === t.key ? C.primary : 'transparent', color: active === t.key ? '#fff' : C.textSecondary, whiteSpace: 'nowrap', transition: 'all 0.15s' }}>
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
    <div style={{ background: C.cardBg, borderRadius: 12, padding: 16, border: `1px solid ${C.cardBorder}` }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <span style={{ fontWeight: 600, color: C.textPrimary, fontSize: 14 }}>{label}</span>
        <button onClick={downloadTemplate} style={{ fontSize: 12, color: C.primary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline' }}>Download Template</button>
      </div>
      <div onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
        onClick={() => fileRef.current?.click()} style={{ border: `2px dashed ${dragOver ? C.primary : C.cardBorder}`, borderRadius: 8, padding: '20px', textAlign: 'center', cursor: 'pointer', background: dragOver ? '#EEF2FF' : C.divider, transition: 'all 0.15s' }}>
        <p style={{ margin: 0, fontSize: 13, color: C.textSecondary }}>Drop CSV here or click to upload</p>
        <input ref={fileRef} type="file" accept=".csv" style={{ display: 'none' }} onChange={e => handleFile(e.target.files?.[0])} />
      </div>
      {status && <p style={{ marginTop: 8, fontSize: 12, color: status.ok ? C.success : C.danger }}>{status.msg}</p>}
    </div>
  );
}

// ─── SETTINGS MODAL ───
function SettingsModal({ open, onClose }) {
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('crm_anthropic_key') || '');
  const [neonConn, setNeonConn] = useState(() => localStorage.getItem('crm_neon_connection') || '');
  const [username, setUsername] = useState(() => localStorage.getItem('crm_username') || '');
  if (!open) return null;
  const save = () => {
    localStorage.setItem('crm_anthropic_key', apiKey);
    localStorage.setItem('crm_neon_connection', neonConn);
    localStorage.setItem('crm_username', username);
    onClose();
  };
  const inputStyle = { width: '100%', padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 13, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' };
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={onClose}>
      <div style={{ background: C.cardBg, borderRadius: 16, padding: 28, width: 460, maxWidth: '90vw', boxShadow: '0 20px 60px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
        <h2 style={{ margin: '0 0 20px', fontSize: 18, fontWeight: 700, color: C.textPrimary }}>Settings</h2>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div>
            <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Anthropic API Key (for AI Import)</label>
            <input type="password" value={apiKey} onChange={e => setApiKey(e.target.value)} placeholder="sk-ant-..." style={inputStyle} />
          </div>
          <div>
            <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Neon Connection String (for Initiatives)</label>
            <input type="password" value={neonConn} onChange={e => setNeonConn(e.target.value)} placeholder="postgresql://user:pass@ep-xxx.region.aws.neon.tech/neondb" style={inputStyle} />
          </div>
          <div>
            <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: C.textSecondary, marginBottom: 4 }}>Your Name (for comments)</label>
            <input type="text" value={username} onChange={e => setUsername(e.target.value)} placeholder="e.g. Alessandro" style={inputStyle} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 24 }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
          <button onClick={save} style={{ padding: '8px 20px', borderRadius: 8, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Save</button>
        </div>
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
    <div style={{ display: 'flex', gap: 2, background: C.divider, borderRadius: 8, padding: 2, width: 'fit-content' }}>
      {opts.map(o => (
        <button key={o} onClick={() => dispatch({ type: 'SET_TAB_PERIOD', tab, period: o })} style={{ padding: '5px 14px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: current === o ? 700 : 500, background: current === o ? C.cardBg : 'transparent', color: current === o ? C.textPrimary : C.textSecondary, cursor: 'pointer', boxShadow: current === o ? '0 1px 3px rgba(0,0,0,0.1)' : 'none', transition: 'all 0.15s' }}>
          {labels[o]}
        </button>
      ))}
    </div>
  );
}

// ─── AI DATA IMPORTER ───
const DATASET_SCHEMAS = {
  emailFlows: { label: 'Email & Flows', fields: ['week','type','flowName','sends','delivered','opens','openRate','clicks','ctr','unsubscribes','unsubRate','revenue','conversions','listSize'], example: DEMO_EMAIL_FLOWS.slice(0, 3) },
  loyalty: { label: 'Loyalty & Rewards', fields: ['month','totalMembers','newEnrollments','pointsIssued','pointsRedeemed','redemptionRate','rewardsRedeemed','rewardCostGBP','revenueFromMembers','revenueFromNonMembers','memberAOV','nonMemberAOV','memberRetentionRate','nonMemberRetentionRate','tierBronze','tierSilver','tierGold','tierPlatinum'], example: DEMO_LOYALTY.slice(0, 2) },
  segments: { label: 'Segments & Lifecycle', fields: ['month','segNew','segActive','segAtRisk','segLapsed','totalCustomers','avgRFMScore','segNewRevenue','segActiveRevenue','segAtRiskRevenue','segLapsedRevenue','migratedAtRiskToActive','migratedActiveToAtRisk','reactivatedFromLapsed','avgOrdersPerActiveCustomer'], example: DEMO_SEGMENTS.slice(0, 2) },
  outreach: { label: 'Direct Outreach', fields: ['week','channel','sends','delivered','responses','responseRate','conversions','conversionRate','revenue','cost'], example: DEMO_OUTREACH.slice(0, 3) },
  beforeAfter: { label: 'Before/After Analysis', fields: ['activity','launchDate','metric','beforeValue','afterValue','beforePeriod','afterPeriod','lift','unit'], example: DEMO_BEFORE_AFTER.slice(0, 2) },
  holdoutTests: { label: 'Holdout Tests', fields: ['testName','testPeriod','sampleSize','controlSize','exposedSize','controlConversionRate','exposedConversionRate','controlRevPerCustomer','exposedRevPerCustomer','incrementalConversionLift','incrementalRevLift','incrementalRevenue','confidence','status'], example: DEMO_HOLDOUT_TESTS.slice(0, 2) },
  activityROI: { label: 'Activity ROI', fields: ['activity','channel','totalCost','attributedRevenue','incrementalRevenue','incrementalROI','customersInfluenced','period'], example: DEMO_ACTIVITY_ROI.slice(0, 2) },
  revenue: { label: 'Revenue', fields: ['week','totalRevenue','subscriptionRevenue','oneTimeRevenue','refunds','netRevenue','totalOrders','aov'], example: DEMO_REVENUE.slice(0, 2) },
  subscriptions: { label: 'Subscriptions', fields: ['month','activeSubscribers','newSubscribers','churnedSubscribers','reactivated','mrr','churnRate','voluntaryChurn','involuntaryChurn','ltv','skipCount'], example: DEMO_SUBSCRIPTIONS.slice(0, 2) },
};

function AIDataImporter({ dispatch, onOpenSettings }) {
  const [selectedDataset, setSelectedDataset] = useState('emailFlows');
  const [rawInput, setRawInput] = useState('');
  const [inputMode, setInputMode] = useState('paste');
  const [loading, setLoading] = useState(false);
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState(null);
  const fileRef = useRef(null);

  const apiKey = localStorage.getItem('crm_anthropic_key');
  const schema = DATASET_SCHEMAS[selectedDataset];

  const handleFileUpload = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => setRawInput(e.target.result);
    reader.readAsText(file);
  };

  const organizeWithAI = async () => {
    if (!apiKey) { setError('No API key configured. Please open Settings.'); return; }
    if (!rawInput.trim()) { setError('Please paste data or upload a file first.'); return; }
    setLoading(true); setError(null); setPreview(null);
    try {
      const resp = await fetch('/api/anthropic/v1/messages', {
        method: 'POST',
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-6-20250514',
          max_tokens: 4096,
          system: `You are a data formatting assistant. Convert the user's raw data into a JSON array matching this exact schema.\n\nFields: ${schema.fields.join(', ')}\n\nExample rows:\n${JSON.stringify(schema.example, null, 2)}\n\nRules:\n- Return ONLY a valid JSON array, no markdown, no explanation\n- Use the exact field names shown\n- Convert dates to the format shown in examples\n- Numeric fields should be numbers, not strings\n- If data is missing a field, use null\n- Ensure all rows have all fields`,
          messages: [{ role: 'user', content: `Convert this data into the ${schema.label} format:\n\n${rawInput}` }],
        }),
      });
      if (!resp.ok) { const err = await resp.json().catch(() => ({})); throw new Error(err.error?.message || `API error ${resp.status}`); }
      const data = await resp.json();
      const text = data.content?.[0]?.text || '';
      const jsonMatch = text.match(/\[[\s\S]*\]/);
      if (!jsonMatch) throw new Error('AI did not return valid JSON array');
      const parsed = JSON.parse(jsonMatch[0]);
      setPreview(parsed);
    } catch (err) {
      setError(err.message);
    } finally { setLoading(false); }
  };

  const importData = () => {
    if (!preview) return;
    dispatch({ type: 'LOAD_DATA', source: selectedDataset, payload: preview });
    setPreview(null); setRawInput(''); setError(null);
  };

  return (
    <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
      <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>AI-Powered Data Import</h3>
      {!apiKey && (
        <div style={{ background: '#FEF3C7', borderRadius: 8, padding: 12, marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 13, color: '#92400E' }}>No API key configured.</span>
          <button onClick={onOpenSettings} style={{ fontSize: 12, color: C.primary, background: 'none', border: 'none', cursor: 'pointer', textDecoration: 'underline', fontWeight: 600 }}>Open Settings</button>
        </div>
      )}
      <div style={{ display: 'flex', gap: 12, marginBottom: 16, alignItems: 'center', flexWrap: 'wrap' }}>
        <select value={selectedDataset} onChange={e => { setSelectedDataset(e.target.value); setPreview(null); setError(null); }} style={{ padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 13, background: C.pageBg, color: C.textPrimary }}>
          {Object.entries(DATASET_SCHEMAS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <div style={{ display: 'flex', gap: 2, background: C.divider, borderRadius: 6, padding: 2 }}>
          {['paste', 'file'].map(m => (
            <button key={m} onClick={() => setInputMode(m)} style={{ padding: '5px 12px', borderRadius: 4, border: 'none', fontSize: 12, fontWeight: inputMode === m ? 700 : 500, background: inputMode === m ? C.cardBg : 'transparent', color: inputMode === m ? C.textPrimary : C.textSecondary, cursor: 'pointer' }}>
              {m === 'paste' ? 'Paste Text' : 'Upload File'}
            </button>
          ))}
        </div>
      </div>
      {inputMode === 'paste' ? (
        <textarea value={rawInput} onChange={e => setRawInput(e.target.value)} placeholder="Paste your raw data here — CSV, tab-separated, JSON, or any format. The AI will organize it..." rows={6} style={{ width: '100%', padding: '10px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 12, fontFamily: 'monospace', resize: 'vertical', background: C.pageBg, color: C.textPrimary, boxSizing: 'border-box' }} />
      ) : (
        <div onClick={() => fileRef.current?.click()} style={{ border: `2px dashed ${C.cardBorder}`, borderRadius: 8, padding: 20, textAlign: 'center', cursor: 'pointer', background: C.divider }}>
          <p style={{ margin: 0, fontSize: 13, color: C.textSecondary }}>{rawInput ? `File loaded (${rawInput.length} chars)` : 'Click to upload CSV, TSV, or text file'}</p>
          <input ref={fileRef} type="file" accept=".csv,.tsv,.txt,.json" style={{ display: 'none' }} onChange={e => handleFileUpload(e.target.files?.[0])} />
        </div>
      )}
      <div style={{ display: 'flex', gap: 10, marginTop: 12, alignItems: 'center' }}>
        <button onClick={organizeWithAI} disabled={loading || !rawInput.trim()} style={{ padding: '8px 20px', borderRadius: 8, border: 'none', background: loading ? C.textTertiary : C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: loading ? 'wait' : 'pointer', opacity: !rawInput.trim() ? 0.5 : 1 }}>
          {loading ? 'Organizing...' : 'Organize with AI'}
        </button>
        {error && <span style={{ fontSize: 12, color: C.danger }}>{error}</span>}
      </div>
      {preview && (
        <div style={{ marginTop: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: C.success }}>{preview.length} rows parsed successfully</span>
            <button onClick={importData} style={{ padding: '8px 20px', borderRadius: 8, border: 'none', background: C.success, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Import Data</button>
          </div>
          <div style={{ maxHeight: 200, overflow: 'auto', borderRadius: 8, border: `1px solid ${C.cardBorder}` }}>
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

// ─── NEON HELPERS ───
async function neonQuery(connectionString, sqlText, params = []) {
  const sql = neon(connectionString);
  return await sql(sqlText, params);
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
}

// ─── INITIATIVES SECTION ───
const STATUS_COLORS = { 'To Do': '#94A3B8', 'In Progress': '#3B82F6', 'Done': '#10B981', 'Blocked': '#EF4444' };
const PRIORITY_COLORS = { 'Low': '#94A3B8', 'Medium': '#3B82F6', 'High': '#F59E0B', 'Urgent': '#EF4444' };
const CATEGORIES = ['Email', 'Loyalty', 'Outreach', 'Segments', 'General'];
const STATUSES = ['To Do', 'In Progress', 'Done', 'Blocked'];
const PRIORITIES = ['Low', 'Medium', 'High', 'Urgent'];

function InitiativesSection({ onOpenSettings }) {
  const connStr = localStorage.getItem('crm_neon_connection');
  const username = localStorage.getItem('crm_username') || 'Anonymous';
  const [initiatives, setInitiatives] = useState([]);
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
  const [dbReady, setDbReady] = useState(false);

  useEffect(() => {
    if (!connStr) return;
    const init = async () => {
      try {
        setLoading(true);
        await initNeonTables(connStr);
        setDbReady(true);
        await loadInitiatives();
      } catch (e) { setError(e.message); } finally { setLoading(false); }
    };
    init();
  }, [connStr]);

  const loadInitiatives = async () => {
    if (!connStr) return;
    try {
      const rows = await neonQuery(connStr, 'SELECT * FROM initiatives ORDER BY CASE priority WHEN \'Urgent\' THEN 0 WHEN \'High\' THEN 1 WHEN \'Medium\' THEN 2 ELSE 3 END, CASE status WHEN \'In Progress\' THEN 0 WHEN \'To Do\' THEN 1 WHEN \'Blocked\' THEN 2 ELSE 3 END, created_at DESC');
      setInitiatives(rows);
    } catch (e) { setError(e.message); }
  };

  const loadComments = async (initiativeId) => {
    if (!connStr) return;
    try {
      const rows = await neonQuery(connStr, 'SELECT * FROM initiative_comments WHERE initiative_id = $1 ORDER BY created_at ASC', [initiativeId]);
      setComments(prev => ({ ...prev, [initiativeId]: rows }));
    } catch (e) { setError(e.message); }
  };

  const saveInitiative = async () => {
    if (!connStr || !formData.title.trim()) return;
    try {
      if (editItem) {
        await neonQuery(connStr, 'UPDATE initiatives SET title=$1, description=$2, status=$3, priority=$4, owner=$5, due_date=$6, category=$7, updated_at=NOW() WHERE id=$8', [formData.title, formData.description, formData.status, formData.priority, formData.owner, formData.due_date || null, formData.category, editItem.id]);
      } else {
        await neonQuery(connStr, 'INSERT INTO initiatives (title, description, status, priority, owner, due_date, category, created_by) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)', [formData.title, formData.description, formData.status, formData.priority, formData.owner, formData.due_date || null, formData.category, username]);
      }
      setShowForm(false); setEditItem(null);
      setFormData({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' });
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const updateStatus = async (id, newStatus) => {
    if (!connStr) return;
    try {
      await neonQuery(connStr, 'UPDATE initiatives SET status=$1, updated_at=NOW() WHERE id=$2', [newStatus, id]);
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const deleteInitiative = async (id) => {
    if (!connStr) return;
    try {
      await neonQuery(connStr, 'DELETE FROM initiatives WHERE id=$1', [id]);
      await loadInitiatives();
    } catch (e) { setError(e.message); }
  };

  const addComment = async (initiativeId) => {
    if (!connStr || !commentText.trim()) return;
    try {
      await neonQuery(connStr, 'INSERT INTO initiative_comments (initiative_id, author, content) VALUES ($1,$2,$3)', [initiativeId, username, commentText]);
      setCommentText('');
      await loadComments(initiativeId);
    } catch (e) { setError(e.message); }
  };

  const toggleExpand = async (id) => {
    if (expandedId === id) { setExpandedId(null); return; }
    setExpandedId(id);
    if (!comments[id]) await loadComments(id);
  };

  if (!connStr) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 16, padding: 60 }}>
        <div style={{ fontSize: 48, opacity: 0.3 }}>&#128203;</div>
        <h3 style={{ margin: 0, fontSize: 18, fontWeight: 600, color: C.textPrimary }}>Initiative Tracker</h3>
        <p style={{ margin: 0, fontSize: 14, color: C.textSecondary, textAlign: 'center', maxWidth: 400 }}>Connect a Neon database to track CRM initiatives, assign owners, and collaborate with comments.</p>
        <button onClick={onOpenSettings} style={{ padding: '10px 24px', borderRadius: 8, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Open Settings</button>
      </div>
    );
  }

  const filtered = initiatives.filter(i => (statusFilter === 'All' || i.status === statusFilter) && (catFilter === 'All' || i.category === catFilter));
  const inputStyle = { padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 13, fontFamily: 'inherit', background: C.pageBg, color: C.textPrimary };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {error && <div style={{ background: '#FEF2F2', borderRadius: 8, padding: 10, fontSize: 12, color: C.danger }}>{error} <button onClick={() => setError(null)} style={{ background: 'none', border: 'none', color: C.danger, cursor: 'pointer', fontWeight: 600, marginLeft: 8 }}>Dismiss</button></div>}

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          {['All', ...STATUSES].map(s => (
            <button key={s} onClick={() => setStatusFilter(s)} style={{ padding: '5px 14px', borderRadius: 20, border: `1px solid ${s === 'All' ? C.cardBorder : STATUS_COLORS[s] || C.cardBorder}`, background: statusFilter === s ? (s === 'All' ? C.textPrimary : STATUS_COLORS[s]) : 'transparent', color: statusFilter === s ? '#fff' : C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
              {s} {s !== 'All' && <span style={{ opacity: 0.7 }}>({initiatives.filter(i => i.status === s).length})</span>}
            </button>
          ))}
          <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ padding: '5px 10px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 12, background: C.pageBg, color: C.textSecondary }}>
            <option value="All">All Categories</option>
            {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <button onClick={() => { setShowForm(true); setEditItem(null); setFormData({ title: '', description: '', status: 'To Do', priority: 'Medium', owner: '', due_date: '', category: 'General' }); }} style={{ padding: '8px 20px', borderRadius: 8, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>+ Add Initiative</button>
      </div>

      {showForm && (
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `2px solid ${C.primary}` }}>
          <h4 style={{ margin: '0 0 16px', fontSize: 14, fontWeight: 600, color: C.textPrimary }}>{editItem ? 'Edit Initiative' : 'New Initiative'}</h4>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
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
            <button onClick={saveInitiative} style={{ padding: '8px 20px', borderRadius: 8, border: 'none', background: C.primary, color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>{editItem ? 'Update' : 'Create'}</button>
            <button onClick={() => { setShowForm(false); setEditItem(null); }} style={{ padding: '8px 20px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      )}

      {loading && <p style={{ textAlign: 'center', color: C.textSecondary, fontSize: 13 }}>Loading...</p>}

      <div style={{ background: C.cardBg, borderRadius: 12, border: `1px solid ${C.cardBorder}`, overflow: 'hidden' }}>
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
                    <select value={item.status} onClick={e => e.stopPropagation()} onChange={e => updateStatus(item.id, e.target.value)} style={{ padding: '3px 8px', borderRadius: 12, border: 'none', background: STATUS_COLORS[item.status], color: '#fff', fontSize: 11, fontWeight: 600, cursor: 'pointer' }}>
                      {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: '10px 12px', fontWeight: 500, color: C.textPrimary }}>{item.title}</td>
                  <td style={{ padding: '10px 12px' }}>
                    <span style={{ padding: '2px 10px', borderRadius: 12, border: `1px solid ${PRIORITY_COLORS[item.priority]}`, color: PRIORITY_COLORS[item.priority], fontSize: 11, fontWeight: 600, background: item.priority === 'Urgent' ? '#FEF2F2' : 'transparent' }}>
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
                        <div key={c.id} style={{ display: 'flex', gap: 8, marginBottom: 8, padding: '8px 10px', background: C.cardBg, borderRadius: 8, border: `1px solid ${C.cardBorder}` }}>
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
                        <input value={commentText} onChange={e => setCommentText(e.target.value)} onKeyDown={e => e.key === 'Enter' && addComment(item.id)} placeholder="Add a comment..." style={{ flex: 1, padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, fontSize: 12, fontFamily: 'inherit', background: C.cardBg, color: C.textPrimary }} />
                        <button onClick={() => addComment(item.id)} disabled={!commentText.trim()} style={{ padding: '8px 16px', borderRadius: 8, border: 'none', background: commentText.trim() ? C.primary : C.textTertiary, color: '#fff', fontSize: 12, fontWeight: 600, cursor: commentText.trim() ? 'pointer' : 'default' }}>Send</button>
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
    </div>
  );
}

// ─── SECTION COMPONENTS ───

function OverviewSection({ state, presentationMode }) {
  const chartHeight = presentationMode ? 400 : 300;
  const latestWeek = [...new Set(state.emailFlows.map(r => r.week))].sort().pop();
  const latestEmailRev = state.emailFlows.filter(r => r.week === latestWeek).reduce((s, r) => s + (r.revenue || 0), 0);
  const latestOutreachWeek = [...new Set(state.outreach.map(r => r.week))].sort().pop();
  const latestOutreachRev = state.outreach.filter(r => r.week === latestOutreachWeek).reduce((s, r) => s + (r.revenue || 0), 0);
  const crmRevWeek = latestEmailRev + latestOutreachRev;
  const latestTotalRev = state.revenue.length > 0 ? state.revenue[state.revenue.length - 1].totalRevenue : 1;
  const crmPct = (crmRevWeek / latestTotalRev) * 100;
  const latestLoyalty = state.loyalty.length > 0 ? state.loyalty[state.loyalty.length - 1] : null;
  const latestSeg = state.segments.length > 0 ? state.segments[state.segments.length - 1] : null;
  const totalCost = state.activityROI.reduce((s, r) => s + r.totalCost, 0);
  const totalIncRev = state.activityROI.reduce((s, r) => s + r.incrementalRevenue, 0);
  const avgROI = totalCost > 0 ? totalIncRev / totalCost : 0;
  const activeTests = state.holdoutTests.filter(t => t.status === 'active').length;

  const allWeeks = [...new Set([...state.emailFlows.map(r => r.week), ...state.outreach.map(r => r.week)])].sort();
  const crmTrend = allWeeks.map(w => {
    const emailRev = state.emailFlows.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0);
    const outreachRev = state.outreach.filter(r => r.week === w).reduce((s, r) => s + (r.revenue || 0), 0);
    const totalRev = state.revenue.find(r => r.week === w)?.totalRevenue || 0;
    return { week: w.slice(5), emailRevenue: emailRev, outreachRevenue: outreachRev, totalRevenue: totalRev };
  });

  const roiSorted = [...state.activityROI].sort((a, b) => b.incrementalRevenue - a.incrementalRevenue);
  const alerts = useMemo(() => generateAlerts(state), [state]);
  const roiColor = (v) => v > 10 ? '#10B981' : v > 3 ? '#3B82F6' : v > 1 ? '#F59E0B' : '#EF4444';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
        <KPICard label="CRM Revenue (Week)" value={crmRevWeek} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="CRM % of Revenue" value={crmPct} format="percent" status={crmPct < 20 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label="Active Loyalty Members" value={latestLoyalty?.totalMembers} format="number" status="good" sparkData={state.loyalty} sparkKey="totalMembers" presentationMode={presentationMode} />
        <KPICard label="At-Risk Customers" value={latestSeg?.segAtRisk} format="number" status={latestSeg?.segAtRisk > 1200 ? 'warning' : 'good'} sparkData={state.segments} sparkKey="segAtRisk" presentationMode={presentationMode} />
        <KPICard label="Avg CRM ROI" value={avgROI} format="multiplier" status={avgROI < 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label="Holdout Tests Running" value={activeTests} format="number" status="good" presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>CRM Revenue Contribution Over Time</h3>
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={crmTrend}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Area type="monotone" dataKey="emailRevenue" stackId="crm" fill="#7C3AED" stroke="#7C3AED" fillOpacity={0.6} name="Email & Flows" />
            <Area type="monotone" dataKey="outreachRevenue" stackId="crm" fill="#25D366" stroke="#25D366" fillOpacity={0.6} name="Outreach" />
            <Line type="monotone" dataKey="totalRevenue" stroke={C.textTertiary} strokeDasharray="5 5" strokeWidth={2} dot={false} name="Total Business Revenue" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Activity Performance Heatmap</h3>
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

        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>CRM Alerts</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {alerts.map((a, i) => (
              <div key={i} style={{ padding: '10px 12px', borderRadius: 8, background: a.severity === 'danger' ? '#FEF2F2' : a.severity === 'warning' ? '#FFFBEB' : '#F0FDF4', borderLeft: `4px solid ${a.severity === 'danger' ? C.danger : a.severity === 'warning' ? C.warning : C.success}` }}>
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
  const period = state.tabPeriods.email || 'weekly';
  const data = period === 'monthly' ? aggregateEmailFlowsByMonth(state.emailFlows) : state.emailFlows;
  const weeks = [...new Set(data.map(r => r.week))].sort();
  const latestWeek = weeks[weeks.length - 1];
  const latestData = data.filter(r => r.week === latestWeek);
  const latestCampaign = latestData.find(r => r.type === 'Campaign');
  const totalEmailRev = latestData.reduce((s, r) => s + (r.revenue || 0), 0);
  const flowRev = latestData.filter(r => r.type === 'Flow').reduce((s, r) => s + (r.revenue || 0), 0);
  const totalRev = state.revenue.length > 0 ? state.revenue[state.revenue.length - 1].totalRevenue : 1;
  const emailPct = (totalEmailRev / totalRev) * 100;

  const weeklyData = weeks.map(w => {
    const rows = data.filter(r => r.week === w);
    const camp = rows.find(r => r.type === 'Campaign');
    const flowR = rows.filter(r => r.type === 'Flow').reduce((s, r) => s + (r.revenue || 0), 0);
    return { week: w.slice(5), campaignRevenue: camp?.revenue || 0, flowRevenue: flowR };
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
      <TimePeriodToggle tab="email" tabPeriods={state.tabPeriods} dispatch={dispatch} />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
        <KPICard label={`Total Email Revenue (${periodLabel})`} value={totalEmailRev} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="Email % of Revenue" value={emailPct} format="percent" status={emailPct < 15 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label="List Size" value={latestCampaign?.listSize} format="number" status="good" presentationMode={presentationMode} />
        <KPICard label="Avg Campaign Open Rate" value={latestCampaign?.openRate} format="percent" status={latestCampaign?.openRate < 35 ? 'danger' : latestCampaign?.openRate < 40 ? 'warning' : 'good'} presentationMode={presentationMode} />
        <KPICard label={`Flow Revenue (${periodLabel})`} value={flowRev} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="Unsubscribe Rate" value={latestCampaign?.unsubRate} format="percent" status={latestCampaign?.unsubRate > 0.5 ? 'danger' : 'good'} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Flow vs Campaign Revenue</h3>
        <ResponsiveContainer width="100%" height={chartHeight}>
          <BarChart data={weeklyData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="campaignRevenue" stackId="a" fill="#7C3AED" name="Campaign" radius={[0,0,0,0]} />
            <Bar dataKey="flowRevenue" stackId="a" fill={C.info} name="Flows" radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Revenue by Flow (All Time)</h3>
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
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Campaign Open Rate & CTR Trends</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Flow Performance (Last 4 Weeks)</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Weekly Email Revenue by Source</h3>
        <ResponsiveContainer width="100%" height={chartHeight}>
          <AreaChart data={waterfallData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} />
            <Tooltip content={<ChartTooltip formatter={(v) => formatCurrency(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {waterfallKeys.map(k => (
              <Area key={k} type="monotone" dataKey={k} stackId="1" fill={CRM_CHANNEL_COLORS[k] || '#7C3AED'} stroke={CRM_CHANNEL_COLORS[k] || '#7C3AED'} fillOpacity={0.7} />
            ))}
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function LoyaltySection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const data = state.loyalty;
  const latest = data.length > 0 ? data[data.length - 1] : null;
  const aovLift = latest ? ((latest.memberAOV - latest.nonMemberAOV) / latest.nonMemberAOV) * 100 : 0;

  const tierData = latest ? [
    { name: 'Bronze', value: latest.tierBronze, fill: TIER_COLORS.Bronze },
    { name: 'Silver', value: latest.tierSilver, fill: TIER_COLORS.Silver },
    { name: 'Gold', value: latest.tierGold, fill: TIER_COLORS.Gold },
    { name: 'Platinum', value: latest.tierPlatinum, fill: TIER_COLORS.Platinum },
  ] : [];

  const chartData = data.map(m => ({
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
  }));

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      <TimePeriodToggle tab="loyalty" tabPeriods={state.tabPeriods} dispatch={dispatch} options={['monthly']} />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
        <KPICard label="Total Members" value={latest?.totalMembers} format="number" status="good" sparkData={data} sparkKey="totalMembers" delta={calcDelta(data, 'totalMembers')} presentationMode={presentationMode} />
        <KPICard label="New Enrollments (Month)" value={latest?.newEnrollments} format="number" status="good" presentationMode={presentationMode} />
        <KPICard label="Redemption Rate" value={latest?.redemptionRate} format="percent" status={latest?.redemptionRate < 10 ? 'warning' : 'good'} sparkData={data} sparkKey="redemptionRate" presentationMode={presentationMode} />
        <KPICard label="Member AOV" value={latest?.memberAOV} format="currencyDecimal" status="good" presentationMode={presentationMode} />
        <KPICard label="Non-Member AOV" value={latest?.nonMemberAOV} format="currencyDecimal" status="neutral" presentationMode={presentationMode} />
        <KPICard label="AOV Lift (Member vs Non)" value={aovLift} format="percent" status="good" presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Loyalty Membership Growth</h3>
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
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Points Economy</h3>
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

        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Tier Distribution</h3>
          <ResponsiveContainer width="100%" height={chartHeight}>
            <PieChart>
              <Pie data={tierData} cx="50%" cy="50%" innerRadius={60} outerRadius={100} paddingAngle={3} dataKey="value" label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>
                {tierData.map((e, i) => <Cell key={i} fill={e.fill} stroke={e.fill} />)}
              </Pie>
              <Tooltip content={<ChartTooltip formatter={(v) => formatNumber(v)} />} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Member vs Non-Member Revenue</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Retention: Members vs Non-Members</h3>
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
  const data = state.segments;
  const latest = data.length > 0 ? data[data.length - 1] : null;

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
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
        <KPICard label="Total Customers" value={latest?.totalCustomers} format="number" status="good" presentationMode={presentationMode} />
        <KPICard label="Active Customers" value={latest?.segActive} format="number" status="good" sparkData={data} sparkKey="segActive" presentationMode={presentationMode} />
        <KPICard label="At-Risk Customers" value={latest?.segAtRisk} format="number" status={latest?.segAtRisk > 1200 ? 'warning' : 'good'} sparkData={data} sparkKey="segAtRisk" presentationMode={presentationMode} />
        <KPICard label="Lapsed Customers" value={latest?.segLapsed} format="number" status="neutral" presentationMode={presentationMode} />
        <KPICard label="Avg RFM Score" value={latest?.avgRFMScore} format="text" status={latest?.avgRFMScore >= 3.0 ? 'good' : 'warning'} presentationMode={presentationMode} />
        <KPICard label="Rescued from At-Risk" value={latest?.migratedAtRiskToActive} format="number" status="good" delta={calcDelta(data, 'migratedAtRiskToActive')} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Lifecycle Segment Distribution</h3>
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

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Segment Revenue Contribution</h3>
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
          <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Migration Matrix (Latest Month)</h3>
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
                          <td key={ci} style={{ padding: '8px', textAlign: 'center', background: val === 0 ? 'transparent' : isPositive ? '#F0FDF4' : isNegative ? '#FEF2F2' : '#F9FAFB', fontWeight: val > 0 ? 600 : 400, color: val === 0 ? C.textTertiary : isPositive ? C.success : isNegative ? C.danger : C.textPrimary, borderRadius: 4 }}>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>At-Risk Recovery Trend</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>RFM Score Trend</h3>
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
    </div>
  );
}

function OutreachSection({ state, presentationMode, dispatch }) {
  const chartHeight = presentationMode ? 400 : 300;
  const period = state.tabPeriods.outreach || 'weekly';
  const data = period === 'monthly' ? aggregateOutreachByMonth(state.outreach) : state.outreach;
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
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
        <KPICard label={`Outreach Revenue (${outPeriodLabel})`} value={latestRev} format="currency" status="good" presentationMode={presentationMode} />
        <KPICard label="WhatsApp Response Rate" value={latestWA?.responseRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="SMS Conversion Rate" value={latestSMS?.conversionRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="Personal Call Conv. Rate" value={latestCall?.conversionRate} format="percent" status="good" presentationMode={presentationMode} />
        <KPICard label="Outreach ROAS" value={outreachROAS} format="multiplier" status={outreachROAS < 3 ? 'warning' : 'good'} presentationMode={presentationMode} />
      </div>

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Outreach Revenue by Channel</h3>
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

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Response Rate by Channel</h3>
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

        <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
          <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Cost per Conversion by Channel</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Channel Performance (Last 4 Weeks)</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>WhatsApp Engagement Trend</h3>
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ComposedChart data={waData}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis dataKey="week" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: C.textTertiary }} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `${v}%`} />
            <Tooltip content={<ChartTooltip formatter={(v, name) => name.includes('Rate') ? formatPercent(v) : formatNumber(v)} />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="sends" fill="#25D366" name="Sends" opacity={0.4} radius={[4,4,0,0]} />
            <Line yAxisId="right" type="monotone" dataKey="responseRate" stroke="#25D366" strokeWidth={2} dot={{ r: 3 }} name="Response Rate" />
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
  const roiColor = (v) => v > 10 ? '#10B981' : v > 3 ? '#3B82F6' : v > 1 ? '#F59E0B' : '#EF4444';

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
    const bg = conf >= 0.90 ? '#D1FAE5' : conf >= 0.80 ? '#DBEAFE' : '#FEF3C7';
    const color = conf >= 0.90 ? '#065F46' : conf >= 0.80 ? '#1E40AF' : '#92400E';
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Before / After Analysis</h3>
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
                      <span style={{ background: r.lift > 50 ? '#D1FAE5' : r.lift > 20 ? '#DBEAFE' : '#FEF3C7', color: r.lift > 50 ? '#065F46' : r.lift > 20 ? '#1E40AF' : '#92400E', padding: '2px 8px', borderRadius: 4, fontWeight: 600, fontSize: 11 }}>+{r.lift.toFixed(1)}%</span>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Control vs Test Group Results</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Incremental Revenue by Test</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Activity ROI Summary</h3>
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

      <div style={{ background: C.cardBg, borderRadius: 12, padding: 20, border: `1px solid ${C.cardBorder}` }}>
        <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Cost vs Incremental Revenue</h3>
        <ResponsiveContainer width="100%" height={chartHeight}>
          <ScatterChart>
            <CartesianGrid strokeDasharray="3 3" stroke={C.divider} />
            <XAxis type="number" dataKey="x" name="Cost" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} label={{ value: 'Total Cost (£)', position: 'bottom', offset: -5, style: { fontSize: 11, fill: C.textSecondary } }} />
            <YAxis type="number" dataKey="y" name="Incremental Revenue" tick={{ fontSize: 11, fill: C.textTertiary }} tickFormatter={v => `£${(v/1000).toFixed(0)}k`} label={{ value: 'Incremental Revenue (£)', angle: -90, position: 'insideLeft', style: { fontSize: 11, fill: C.textSecondary } }} />
            <Tooltip content={({ active, payload }) => {
              if (!active || !payload?.length) return null;
              const d = payload[0].payload;
              return (
                <div style={{ background: C.cardBg, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: '10px 14px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
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
    </div>
  );
}

function DataImportSection({ state, dispatch, onOpenSettings }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <AIDataImporter dispatch={dispatch} onOpenSettings={onOpenSettings} />
      <h3 style={{ margin: '12px 0 0', fontSize: 15, fontWeight: 600, color: C.textPrimary }}>Manual CSV Upload</h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: 16 }}>
        <CSVUploader label="Email & Flows Data" source="emailFlows" requiredHeaders={['week','type','sends','revenue']} dispatch={dispatch} />
        <CSVUploader label="Loyalty & Rewards Data" source="loyalty" requiredHeaders={['month','totalMembers','memberAOV']} dispatch={dispatch} />
        <CSVUploader label="Customer Segments Data" source="segments" requiredHeaders={['month','segNew','segActive','segAtRisk','segLapsed']} dispatch={dispatch} />
        <CSVUploader label="Direct Outreach Data" source="outreach" requiredHeaders={['week','channel','sends','revenue','cost']} dispatch={dispatch} />
        <CSVUploader label="Before/After Analysis" source="beforeAfter" requiredHeaders={['activity','metric','beforeValue','afterValue','lift']} dispatch={dispatch} />
        <CSVUploader label="Holdout Test Results" source="holdoutTests" requiredHeaders={['testName','controlConversionRate','exposedConversionRate','incrementalRevenue']} dispatch={dispatch} />
        <CSVUploader label="Activity ROI Data" source="activityROI" requiredHeaders={['activity','totalCost','incrementalRevenue','incrementalROI']} dispatch={dispatch} />
        <CSVUploader label="Revenue Data" source="revenue" requiredHeaders={['week','totalRevenue','netRevenue']} dispatch={dispatch} />
        <CSVUploader label="Subscription Data" source="subscriptions" requiredHeaders={['month','activeSubscribers','mrr','churnRate']} dispatch={dispatch} />
      </div>
      <div style={{ display: 'flex', gap: 12, justifyContent: 'center', paddingTop: 12 }}>
        <button onClick={() => dispatch({ type: 'RESET_DEMO' })} style={{ padding: '10px 24px', borderRadius: 8, border: `1px solid ${C.primary}`, background: 'transparent', color: C.primary, fontWeight: 600, cursor: 'pointer', fontSize: 13 }}>Reset to Demo Data</button>
        <button onClick={() => dispatch({ type: 'CLEAR_ALL' })} style={{ padding: '10px 24px', borderRadius: 8, border: `1px solid ${C.danger}`, background: 'transparent', color: C.danger, fontWeight: 600, cursor: 'pointer', fontSize: 13 }}>Clear All Data</button>
      </div>
    </div>
  );
}

// ─── MAIN APP ───
export default function App() {
  const [state, dispatch] = useReducer(reducer, initialState);
  const [presentationMode, setPresentationMode] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const openSettings = useCallback(() => setSettingsOpen(true), []);

  return (
    <div style={{ minHeight: '100vh', background: C.pageBg, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }}>
      <SettingsModal open={settingsOpen} onClose={() => setSettingsOpen(false)} />
      <header style={{ background: C.cardBg, borderBottom: `1px solid ${C.cardBorder}`, padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h1 style={{ margin: 0, fontSize: presentationMode ? 24 : 20, fontWeight: 700, color: C.textPrimary }}>CRM Manager Dashboard</h1>
          <p style={{ margin: '2px 0 0', fontSize: 13, color: C.textSecondary }}>Subscription DTC — CRM Performance & Incrementality</p>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <button onClick={openSettings} title="Settings" style={{ padding: '8px 12px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, background: 'transparent', color: C.textSecondary, fontSize: 16, cursor: 'pointer', lineHeight: 1 }}>&#9881;</button>
          <button onClick={() => setPresentationMode(!presentationMode)} style={{ padding: '8px 16px', borderRadius: 8, border: `1px solid ${C.cardBorder}`, background: presentationMode ? C.primary : 'transparent', color: presentationMode ? '#fff' : C.textSecondary, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
            {presentationMode ? 'Exit Presentation' : 'Presentation Mode'}
          </button>
        </div>
      </header>

      <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 24px 40px' }}>
        <TabNav tabs={TABS} active={state.activeTab} onSelect={t => dispatch({ type: 'SET_TAB', payload: t })} />
        <div style={{ paddingTop: 20 }}>
          {state.activeTab === 'overview' && <OverviewSection state={state} presentationMode={presentationMode} />}
          {state.activeTab === 'email' && <EmailFlowsSection state={state} presentationMode={presentationMode} dispatch={dispatch} />}
          {state.activeTab === 'loyalty' && <LoyaltySection state={state} presentationMode={presentationMode} dispatch={dispatch} />}
          {state.activeTab === 'segments' && <SegmentsSection state={state} presentationMode={presentationMode} dispatch={dispatch} />}
          {state.activeTab === 'outreach' && <OutreachSection state={state} presentationMode={presentationMode} dispatch={dispatch} />}
          {state.activeTab === 'incrementality' && <IncrementalitySection state={state} presentationMode={presentationMode} />}
          {state.activeTab === 'initiatives' && <InitiativesSection onOpenSettings={openSettings} />}
          {state.activeTab === 'import' && <DataImportSection state={state} dispatch={dispatch} onOpenSettings={openSettings} />}
        </div>
      </div>
    </div>
  );
}
