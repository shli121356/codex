import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  Area,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  Legend,
  Line,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from 'recharts';
import {
  Activity,
  BarChart3,
  CalendarClock,
  CheckCircle2,
  FileSpreadsheet,
  RefreshCw,
  Search,
  ShoppingBag,
  SlidersHorizontal,
  Upload
} from 'lucide-react';
import { readSheet } from 'read-excel-file/browser';
import './styles.css';

const SAMPLE_FILE = '/sample-shop.xlsx';
const WATCHED_WORKBOOKS_API = '/api/watched-workbooks';
const WATCHED_WORKBOOKS_EVENTS = '/api/watched-workbooks/events';
const ACCENTS = ['#10d2a0', '#7fb0d2', '#b8cdf4', '#f2c16b', '#ee6b73', '#8b7cf6'];
const PRODUCT_COLORS = ['#10d2a0', '#7fb0d2', '#f2c16b', '#ee6b73', '#8b7cf6', '#38d6e5', '#ff9f43', '#c6d6ff'];
const STATUS_COLORS = {
  已关闭: '#18d4a6',
  已完成: '#66a9e6',
  未知状态: '#f2c16b'
};
const COMPLETED = '已完成';
const GIFT_PRODUCT_NAME = '400%娃娃鞋';
const RESPONSIVE_CHART_PROPS = {
  width: '100%',
  height: '100%',
  minWidth: 0,
  minHeight: 0,
  initialDimension: { width: 1, height: 1 }
};

function emptyDashboardData() {
  return { orders: [], productSummary: [], homeMetrics: {}, homeMetricsByStore: {}, stores: [], fileNames: [] };
}

function normalizeMoney(value) {
  const number = String(value ?? '').replace(/[^\d.-]/g, '');
  return Number(number || 0);
}

function normalizeCount(value) {
  const number = String(value ?? '').replace(/[^\d.-]/g, '');
  return Number(number || 0);
}

function cleanHeader(value) {
  return String(value ?? '').replace(/_x000D_/g, '').replace(/\s+/g, '');
}

function fileBaseName(name) {
  return String(name || '未命名店铺').replace(/\.(xlsx|xls)$/i, '');
}

function dashboardTitleFromFiles(fileNames) {
  const names = (fileNames || []).map(fileBaseName).filter(Boolean);
  if (!names.length) return '店铺数据看板';
  if (names.length === 1) return `${names[0]}数据看板`;
  return `${names.slice(0, 3).join(' + ')}${names.length > 3 ? ` 等${names.length}店` : ''}数据看板`;
}

function dashboardSummaryTitleFromFiles(fileNames) {
  const names = (fileNames || []).map(fileBaseName).filter(Boolean);
  if (!names.length) return '全部店铺汇总看板';
  if (names.length === 1) return `${names[0]}数据看板`;
  return `${names.slice(0, 3).join(' + ')}${names.length > 3 ? ` 等${names.length}店` : ''}汇总看板`;
}

function parseMetricGrid(rows) {
  const metrics = {};
  const filledRows = rows.filter(row => row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));
  for (let i = 0; i < filledRows.length; i += 2) {
    const keys = filledRows[i] || [];
    const values = filledRows[i + 1] || [];
    keys.forEach((key, index) => {
      const name = cleanHeader(key);
      if (name) metrics[name] = values[index] ?? '';
    });
  }
  return metrics;
}

async function readWorkbookSheet(file, sheet) {
  try {
    return await readSheet(file, sheet);
  } catch {
    return [];
  }
}

async function parseWorkbook(file, sourceName = file?.name || '店铺1.xlsx') {
  const store = fileBaseName(sourceName);
  const orderRows = await readWorkbookSheet(file, 'Sheet1');
  const fallbackRows = orderRows.length ? orderRows : await readSheet(file);
  const summaryRows = await readWorkbookSheet(file, '商品价格总览');
  const homeRows = await readWorkbookSheet(file, '主页数据');
  const homeMetrics = parseMetricGrid(homeRows);

  const orders = fallbackRows.slice(1).map((row, index) => {
    const rawDate = String(row[3] ?? '').replace(/^下单时间\s*/, '').trim();
    const date = rawDate.match(/\d{4}-\d{2}-\d{2}/)?.[0] || '';
    return {
      id: String(row[2] ?? '').replace(/^订单编号\s*/, '').trim() || `ORDER-${index + 1}`,
      product: String(row[0] || '').trim() || '未命名商品',
      sku: String(row[1] || '').trim() || '未标注规格',
      date,
      month: date.slice(0, 7) || '未知月份',
      time: rawDate,
      price: normalizeMoney(row[4]),
      status: String(row[5] || '未知状态').trim(),
      store
    };
  }).filter(order => order.id || order.product || order.price);

  const summaryHeaders = (summaryRows[0] || []).map(cleanHeader);
  const productSummary = summaryRows.slice(1).filter(row => row.some(Boolean)).map(row => ({
    product: String(row[0] || '未命名商品').trim(),
    refundQty: normalizeCount(row[summaryHeaders.indexOf('商品退款数量')] ?? row[1]),
    soldQty: normalizeCount(row[summaryHeaders.indexOf('商品销售数量')] ?? row[2]),
    amount: normalizeMoney(row[summaryHeaders.indexOf('商品成交总价')] ?? row[3]),
    store
  }));

  return {
    orders,
    productSummary,
    homeMetrics,
    homeMetricsByStore: { [store]: homeMetrics },
    stores: [store],
    fileNames: [sourceName]
  };
}

function mergeMetrics(metricsList) {
  return metricsList.reduce((merged, metrics) => {
    Object.entries(metrics || {}).forEach(([key, value]) => {
      if (value === '-' || value === '' || value === null || value === undefined) {
        merged[key] ??= value;
        return;
      }
      merged[key] = normalizeCount(merged[key]) + normalizeCount(value);
    });
    return merged;
  }, {});
}

function mergeProductSummary(summaries) {
  const grouped = new Map();
  summaries.forEach(item => {
    const current = grouped.get(item.product) || { product: item.product, refundQty: 0, soldQty: 0, amount: 0 };
    current.refundQty += item.refundQty || 0;
    current.soldQty += item.soldQty || 0;
    current.amount += item.amount || 0;
    grouped.set(item.product, current);
  });
  return [...grouped.values()];
}

function mergeWorkbookData(workbooks) {
  const homeMetricsByStore = workbooks.reduce((acc, workbook) => {
    (workbook.stores || []).forEach(store => {
      acc[store] = workbook.homeMetricsByStore?.[store] || workbook.homeMetrics || {};
    });
    return acc;
  }, {});

  return {
    orders: workbooks.flatMap(workbook => workbook.orders || []),
    productSummary: workbooks.flatMap(workbook => workbook.productSummary || []),
    homeMetrics: mergeMetrics(workbooks.map(workbook => workbook.homeMetrics)),
    homeMetricsByStore,
    stores: [...new Set(workbooks.flatMap(workbook => workbook.stores || []))],
    fileNames: workbooks.flatMap(workbook => workbook.fileNames || [])
  };
}

function groupBy(items, keyFn) {
  return items.reduce((acc, item) => {
    const key = keyFn(item);
    acc[key] ||= [];
    acc[key].push(item);
    return acc;
  }, {});
}

function sum(items, keyFn) {
  return items.reduce((total, item) => total + keyFn(item), 0);
}

function formatCurrency(value) {
  return new Intl.NumberFormat('zh-CN', {
    style: 'currency',
    currency: 'CNY',
    maximumFractionDigits: value >= 1000 ? 0 : 2
  }).format(value || 0);
}

function formatNumber(value) {
  return new Intl.NumberFormat('zh-CN').format(value || 0);
}

const PRODUCT_ALIAS_RULES = [
  ['娃三岁400%', '400%赠品'],
  ['400%娃娃鞋', '400%赠品'],
  ['娃娃鞋一只', '单只娃鞋'],
  ['需随单带走', '单只娃鞋'],
  ['娃娃鞋', '娃娃鞋款'],
  ['童话睡前故事', '童话毛绒'],
  ['毛绒盲盒', '童话毛绒'],
  ['贝壳拆拆乐', '香片冰块'],
  ['香片冰块', '香片冰块'],
  ['可爱萌粒', '萌粒摆件'],
  ['diy女生桌面', '萌粒摆件'],
  ['随机萌粒', '随机萌粒'],
  ['流沙萌粒', '流沙萌粒'],
  ['大流沙', '大流沙款']
];

function productAlias(name) {
  const raw = String(name || '').trim();
  if (!raw || raw === '未命名商品') return '未命名品';

  const matched = PRODUCT_ALIAS_RULES.find(([keyword]) => raw.includes(keyword));
  if (matched) return matched[1];

  const withoutBrand = raw.replace(/【[^】]+】/g, '');
  return withoutBrand
    .replace(/diy/gi, 'DIY')
    .replace(/潮流摆件|超萌好物|女生|桌面|生日礼物|车载装饰|创意手作|饰品|可爱卡通公仔|儿童|娃三岁/g, '')
    .slice(0, 6) || '其他商品';
}

function isGiftProduct(name) {
  const raw = String(name || '');
  return raw.includes('娃三岁400%') || raw.includes('400%娃娃鞋');
}

function chartAlias(name) {
  const raw = String(name || '').trim();
  if (!raw || raw === '未命名商品') return '未命名品';
  if (raw.includes('娃三岁400%') || raw.includes('400%娃娃鞋')) return '400%赠品';
  if (raw.includes('娃娃鞋一只') || raw.includes('需随单带走')) return '单只娃鞋';
  if (raw.includes('娃娃鞋')) return '娃娃鞋款';
  if (raw.includes('童话睡前故事') || raw.includes('毛绒盲盒')) return '童话毛绒';
  if (raw.includes('贝壳拆拆乐') || raw.includes('香片冰块')) return '香片冰块';
  if (raw.includes('可爱萌粒') || raw.toLowerCase().includes('diy女生桌面')) return '萌粒摆件';
  if (raw.includes('流沙萌粒')) return '流沙萌粒';
  if (raw.includes('随机萌粒')) return '随机萌粒';
  if (raw.includes('大流沙')) return '大流沙款';
  return productAlias(raw).slice(0, 4);
}

function getTopProducts(orders) {
  return Object.entries(groupBy(orders.filter(order => !isGiftProduct(order.product)), order => order.product))
    .map(([product, list]) => ({
      product,
      amount: sum(list, order => order.price),
      orders: list.length
    }))
    .sort((a, b) => b.amount - a.amount || b.orders - a.orders)
    .slice(0, 6)
    .map(item => item.product);
}

function scopedHomeMetrics(data, selectedStores) {
  if (!selectedStores?.length) return data.homeMetrics || {};
  const metricsByStore = data.homeMetricsByStore || {};
  return mergeMetrics(selectedStores.map(store => metricsByStore[store] || {}));
}

function calcDashboard(data, filters) {
  const allOrders = data.orders || [];
  const selectedProducts = filters.products;
  const selectedStatuses = filters.statuses;
  const selectedStores = filters.stores;
  const orders = allOrders.filter(order => {
    const statusOk = !selectedStatuses.length || selectedStatuses.includes(order.status);
    const productOk = !selectedProducts.length || selectedProducts.includes(order.product);
    const storeOk = !selectedStores.length || selectedStores.includes(order.store);
    return statusOk && productOk && storeOk;
  });
  const completedOrders = orders.filter(order => order.status === COMPLETED);
  const riskOrders = orders.filter(order => order.status !== COMPLETED);
  const coreProductOrders = orders.filter(order => !isGiftProduct(order.product));
  const productGroups = groupBy(coreProductOrders, order => order.product);
  const statusGroups = groupBy(orders, order => order.status);
  const months = Object.keys(groupBy(orders, order => order.month)).sort();

  const productBars = Object.entries(productGroups)
    .map(([product, list]) => ({
      product,
      label: chartAlias(product),
      displayName: productAlias(product),
      amount: sum(list, order => order.price),
      orders: list.length,
      completed: sum(list.filter(order => order.status === COMPLETED), order => order.price),
      risk: sum(list.filter(order => order.status !== COMPLETED), order => order.price)
    }))
    .sort((a, b) => b.amount - a.amount)
    .map((item, index) => ({ ...item, color: PRODUCT_COLORS[index % PRODUCT_COLORS.length] }));

  const monthly = months.map(month => {
    const list = orders.filter(order => order.month === month);
    return {
      month,
      amount: sum(list, order => order.price),
      completed: sum(list.filter(order => order.status === COMPLETED), order => order.price),
      risk: sum(list.filter(order => order.status !== COMPLETED), order => order.price),
      orders: list.length
    };
  });

  const statusPie = Object.entries(statusGroups).map(([name, list], index) => ({
    name,
    value: list.length,
    amount: sum(list, order => order.price),
    color: STATUS_COLORS[name] || ACCENTS[index % ACCENTS.length]
  }));

  const relevantSummaries = (data.productSummary || []).filter(item => !selectedStores.length || selectedStores.includes(item.store));
  const summaryMap = new Map(mergeProductSummary(relevantSummaries).map(item => [item.product, item]));
  const giftOrders = orders.filter(order => isGiftProduct(order.product));
  const giftSummary = relevantSummaries
    .filter(item => isGiftProduct(item.product))
    .reduce((acc, item) => ({
      soldQty: acc.soldQty + (item.soldQty || 0),
      refundQty: acc.refundQty + (item.refundQty || 0)
    }), { soldQty: 0, refundQty: 0 });
  const giftQty = giftSummary.soldQty || giftOrders.length;
  const productMix = productBars.map(item => {
    const summary = summaryMap.get(item.product);
    const isDollShoe = productAlias(item.product) === '娃娃鞋款';
    return {
      ...item,
      soldQty: summary?.soldQty ?? item.orders,
      refundQty: summary?.refundQty ?? 0,
      refundRate: summary?.soldQty ? (summary.refundQty / summary.soldQty) * 100 : 0,
      giftNote: isDollShoe && giftQty ? `赠品：${GIFT_PRODUCT_NAME} x${formatNumber(giftQty)}` : ''
    };
  }).slice(0, 8);

  const firstDate = orders.map(order => order.date).filter(Boolean).sort()[0] || '';
  const lastDate = orders.map(order => order.date).filter(Boolean).sort().at(-1) || '';
  const totalAmount = sum(orders, order => order.price);
  const completedAmount = sum(completedOrders, order => order.price);
  const riskAmount = sum(riskOrders, order => order.price);
  const averageOrder = completedOrders.length ? completedAmount / completedOrders.length : 0;
  const home = scopedHomeMetrics(data, selectedStores);

  return {
    orders,
    productBars,
    productMix,
    monthly,
    statusPie,
    kpis: {
      totalAmount,
      completedAmount,
      riskAmount,
      averageOrder,
      orderCount: orders.length,
      completedCount: completedOrders.length,
      closeCount: riskOrders.length,
      exposure: normalizeCount(home['商品曝光人数']),
      clicks: normalizeCount(home['商品点击人数'])
    },
    homeMetrics: home,
    period: `${firstDate || '未识别'} 至 ${lastDate || '未识别'}`
  };
}

function FilterChip({ active, children, onClick, title }) {
  return (
    <button className={`chip ${active ? 'active' : ''}`} onClick={onClick} title={title}>
      {children}
    </button>
  );
}

function KpiCard({ title, value, sub, icon: Icon, tone = 'green' }) {
  return (
    <section className={`panel kpi-card ${tone}`}>
      <div className="kpi-top">
        <span>{title}</span>
        <Icon size={18} />
      </div>
      <strong>{value}</strong>
      <small>{sub}</small>
    </section>
  );
}

function ChartPanel({ title, note, children, className = '' }) {
  return (
    <section className={`panel chart-panel ${className}`}>
      <div className="panel-heading">
        <div>
          <h2>{title}</h2>
          {note ? <p>{note}</p> : null}
        </div>
        <BarChart3 size={17} />
      </div>
      <div className="chart-body">{children}</div>
    </section>
  );
}

function MiniTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="tooltip">
      <b>{label}</b>
      {payload.map(item => (
        <span key={item.dataKey}>
          {item.name}: {item.dataKey === 'orders' ? `${formatNumber(item.value)} 笔` : typeof item.value === 'number' ? formatCurrency(item.value) : item.value}
        </span>
      ))}
    </div>
  );
}

function ProductAxisTick({ x, y, payload }) {
  const text = String(payload?.value || '');
  const firstLine = text.slice(0, Math.ceil(text.length / 2));
  const secondLine = text.slice(Math.ceil(text.length / 2));
  return (
    <text x={x} y={y} textAnchor="middle" fill="#d7e4e7" fontSize={12} fontWeight={800}>
      <tspan x={x} dy="0">{firstLine}</tspan>
      {secondLine ? <tspan x={x} dy="15">{secondLine}</tspan> : null}
    </text>
  );
}

function App() {
  const [data, setData] = useState(emptyDashboardData());
  const [fileNames, setFileNames] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [filters, setFilters] = useState({ statuses: [], products: [], stores: [] });
  const [storeMenuOpen, setStoreMenuOpen] = useState(false);
  const [folderStatus, setFolderStatus] = useState('正在连接表格文件夹...');
  const storeSelectRef = useRef(null);

  useEffect(() => {
    let cancelled = false;

    async function applyWorkbooks(workbooks, statusLabel) {
      if (!Array.isArray(workbooks)) return false;
      if (!workbooks.length) {
        if (!cancelled) {
          setData(emptyDashboardData());
          setFileNames([]);
          setFilters({ statuses: [], products: [], stores: [] });
          setFolderStatus(`${statusLabel}：识别到 0 个表格`);
          setLoading(false);
        }
        return true;
      }

      if (!cancelled) {
        setLoading(true);
        setError('');
        setFolderStatus(`${statusLabel}：识别到 ${workbooks.length} 个表格`);
      }

      try {
        const parsedWorkbooks = await Promise.all(workbooks.map(async workbook => {
          const response = await fetch(workbook.url, { cache: 'no-store' });
          if (!response.ok) throw new Error(`读取失败：${workbook.name}`);
          const blob = await response.blob();
          return parseWorkbook(blob, workbook.name);
        }));
        if (cancelled) return true;
        const merged = mergeWorkbookData(parsedWorkbooks);
        setData(merged);
        setFileNames(merged.fileNames);
        setFilters({ statuses: [], products: [], stores: merged.stores.length > 1 ? [merged.stores[0]] : [] });
        setStoreMenuOpen(false);
        setFolderStatus(`${statusLabel}：已载入 ${merged.fileNames.length} 个表格`);
        return true;
      } catch (err) {
        if (!cancelled) {
          setError('监听文件夹中的表格读取失败，请确认文件没有被 Excel 占用，或使用右上角上传表格。');
          setFolderStatus(`${statusLabel}：读取失败`);
        }
        return false;
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    async function loadSampleFallback() {
      try {
        const response = await fetch(SAMPLE_FILE);
        const blob = await response.blob();
        const parsed = await parseWorkbook(blob, '店铺1.xlsx');
        if (cancelled) return;
        setData(parsed);
        setFileNames(parsed.fileNames);
        setFilters(current => ({ ...current, stores: parsed.stores?.length > 1 ? [parsed.stores[0]] : [] }));
        setFolderStatus('文件夹监听未连接，当前显示示例表格');
      } catch {
        if (!cancelled) setError('默认表格加载失败，请使用右上角上传 Excel。');
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    async function loadInitialFolder() {
      try {
        const response = await fetch(WATCHED_WORKBOOKS_API, { cache: 'no-store' });
        if (!response.ok) throw new Error('watch api unavailable');
        const payload = await response.json();
        await applyWorkbooks(payload.files, `监听 ${payload.folder}`);
      } catch {
        await loadSampleFallback();
      }
    }

    loadInitialFolder();

    const events = new EventSource(WATCHED_WORKBOOKS_EVENTS);
    events.addEventListener('workbooks', event => {
      try {
        const payload = JSON.parse(event.data);
        applyWorkbooks(payload.files, `监听 ${payload.folder}`);
      } catch {
        setFolderStatus('文件夹监听事件解析失败');
      }
    });
    events.onerror = () => {
      setFolderStatus('文件夹监听连接中断，手动上传仍可使用');
    };

    return () => {
      cancelled = true;
      events.close();
    };
  }, []);

  const storeOptions = useMemo(() => data.stores || [], [data]);
  const activeStore = filters.stores.length === 1 ? filters.stores[0] : '';
  const pageTitle = useMemo(() => {
    if (activeStore) return `${activeStore}数据看板`;
    return storeOptions.length > 1 ? dashboardSummaryTitleFromFiles(fileNames) : dashboardTitleFromFiles(fileNames);
  }, [activeStore, fileNames, storeOptions.length]);
  const fileLabel = fileNames.length > 2 ? `${fileNames.slice(0, 2).join('、')} 等${fileNames.length}个表格` : fileNames.join('、');
  const storeFilteredOrders = useMemo(() => {
    return (data.orders || []).filter(order => !filters.stores.length || filters.stores.includes(order.store));
  }, [data.orders, filters.stores]);
  const statusOptions = useMemo(() => [...new Set(storeFilteredOrders.map(order => order.status))], [storeFilteredOrders]);
  const productOptions = useMemo(() => getTopProducts(storeFilteredOrders), [storeFilteredOrders]);
  const dashboard = useMemo(() => calcDashboard(data, filters), [data, filters]);
  const storeTabs = useMemo(() => {
    return storeOptions.map(store => {
      const orders = (data.orders || []).filter(order => order.store === store);
      return {
        store,
        orders: orders.length,
        amount: sum(orders, order => order.price)
      };
    });
  }, [data.orders, storeOptions]);
  const totalStoreTab = useMemo(() => ({
    orders: (data.orders || []).length,
    amount: sum(data.orders || [], order => order.price)
  }), [data.orders]);
  const activeStoreStats = useMemo(() => {
    if (!activeStore) {
      return { name: '全部店铺汇总', orders: totalStoreTab.orders, amount: totalStoreTab.amount };
    }
    const current = storeTabs.find(item => item.store === activeStore);
    return {
      name: activeStore,
      orders: current?.orders || 0,
      amount: current?.amount || 0
    };
  }, [activeStore, storeTabs, totalStoreTab]);
  const storeSelectOptions = useMemo(() => [
    ...storeTabs.map(item => ({
      value: item.store,
      label: item.store,
      type: '店铺页面',
      orders: item.orders,
      amount: item.amount
    })),
    {
      value: '__summary__',
      label: '全部店铺汇总',
      type: '汇总视图',
      orders: totalStoreTab.orders,
      amount: totalStoreTab.amount
    }
  ], [storeTabs, totalStoreTab]);

  useEffect(() => {
    document.title = pageTitle;
  }, [pageTitle]);

  useEffect(() => {
    function handlePointerDown(event) {
      if (!storeSelectRef.current?.contains(event.target)) {
        setStoreMenuOpen(false);
      }
    }

    function handleKeyDown(event) {
      if (event.key === 'Escape') setStoreMenuOpen(false);
    }

    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  async function handleFile(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) return;
    setLoading(true);
    setError('');
    try {
      const parsedWorkbooks = await Promise.all(files.map(file => parseWorkbook(file, file.name)));
      const merged = mergeWorkbookData(parsedWorkbooks);
      setData(merged);
      setFileNames(merged.fileNames);
      setFilters({ statuses: [], products: [], stores: merged.stores.length > 1 ? [merged.stores[0]] : [] });
      setStoreMenuOpen(false);
      setFolderStatus(`手动上传：已载入 ${merged.fileNames.length} 个表格`);
    } catch (err) {
      setError('解析失败，请确认每个工作簿都包含 Sheet1、商品价格总览、主页数据等字段。');
    } finally {
      setLoading(false);
      event.target.value = '';
    }
  }

  function toggleFilter(type, value) {
    setFilters(current => {
      const values = current[type];
      return {
        ...current,
        [type]: values.includes(value) ? values.filter(item => item !== value) : [...values, value]
      };
    });
  }

  function openStorePage(store) {
    setFilters({ statuses: [], products: [], stores: store ? [store] : [] });
    setStoreMenuOpen(false);
  }

  const maxProductAmount = Math.max(...dashboard.productMix.map(item => item.amount), 1);
  const sourceLabel = fileLabel || folderStatus;

  return (
    <main className="dashboard">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><ShoppingBag size={22} /></div>
          <div>
            <h1>{pageTitle}</h1>
            <span>{activeStore ? `当前店铺：${activeStore}` : '当前视图：全部店铺汇总'} · {sourceLabel} · {dashboard.period}</span>
          </div>
        </div>
        <div className="toolbar">
          <div className="search-box">
            <Search size={16} />
            <span>{formatNumber(dashboard.kpis.orderCount)} 条订单</span>
          </div>
          <div className="monitor-box" title={folderStatus}>
            <FileSpreadsheet size={16} />
            <span>{folderStatus}</span>
          </div>
          <label className="upload-btn">
            <Upload size={16} />
            上传多个表格
            <input type="file" accept=".xlsx,.xls" multiple onChange={handleFile} />
          </label>
        </div>
      </header>

      {storeTabs.length > 1 ? (
        <nav className="store-pages" aria-label="店铺页面切换">
          <div className="store-pages-title">
            <span>店铺页面</span>
            <b>{activeStore || '全部汇总'}</b>
          </div>
          <div className="store-select-panel" ref={storeSelectRef}>
            <span>选择店铺看板</span>
            <div className={`store-select-frame ${storeMenuOpen ? 'open' : ''}`}>
              <button
                className="store-select-trigger"
                type="button"
                aria-haspopup="listbox"
                aria-expanded={storeMenuOpen}
                onClick={() => setStoreMenuOpen(open => !open)}
              >
                <span>
                  <b>{activeStoreStats.name}</b>
                  <small>{formatNumber(activeStoreStats.orders)} 单 · {formatCurrency(activeStoreStats.amount)}</small>
                </span>
                <i />
              </button>
              <div className="store-select-menu" role="listbox">
                {storeSelectOptions.map(option => {
                  const selected = (activeStore || '__summary__') === option.value;
                  return (
                    <button
                      className={`store-option ${selected ? 'active' : ''}`}
                      type="button"
                      role="option"
                      aria-selected={selected}
                      key={option.value}
                      onClick={() => openStorePage(option.value === '__summary__' ? '' : option.value)}
                    >
                      <span>
                        <em>{option.type}</em>
                        <b>{option.label}</b>
                      </span>
                      <strong>{formatNumber(option.orders)} 单 · {formatCurrency(option.amount)}</strong>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
          <div className="store-select-stats">
            <span>当前视图</span>
            <b>{activeStoreStats.name}</b>
            <small>{formatNumber(activeStoreStats.orders)} 单 · {formatCurrency(activeStoreStats.amount)}</small>
          </div>
        </nav>
      ) : null}

      <nav className="filter-row">
        <div className="filter-group">
          <span>订单状态</span>
          <FilterChip active={!filters.statuses.length} onClick={() => setFilters(v => ({ ...v, statuses: [] }))}>全部</FilterChip>
          {statusOptions.map(status => (
            <FilterChip key={status} active={filters.statuses.includes(status)} onClick={() => toggleFilter('statuses', status)}>
              {status}
            </FilterChip>
          ))}
        </div>
        <div className="filter-group product-filter">
          <span>核心商品</span>
          <FilterChip active={!filters.products.length} onClick={() => setFilters(v => ({ ...v, products: [] }))}>全部</FilterChip>
          {productOptions.map(product => (
            <FilterChip key={product} active={filters.products.includes(product)} onClick={() => toggleFilter('products', product)} title={product}>
              {productAlias(product)}
            </FilterChip>
          ))}
        </div>
        <button className="icon-btn" onClick={() => setFilters(v => ({ ...v, statuses: [], products: [] }))} title="重置筛选">
          <RefreshCw size={16} />
        </button>
      </nav>

      {error ? <div className="alert">{error}</div> : null}
      {loading ? <div className="loading">正在生成看板...</div> : null}

      <section className="kpi-grid">
        <KpiCard title="总订单金额" value={formatCurrency(dashboard.kpis.totalAmount)} sub={`订单 ${formatNumber(dashboard.kpis.orderCount)} 笔`} icon={Activity} />
        <KpiCard title="有效成交金额" value={formatCurrency(dashboard.kpis.completedAmount)} sub={`已完成 ${formatNumber(dashboard.kpis.completedCount)} 笔`} icon={CheckCircle2} tone="blue" />
        <KpiCard title="关闭/退款风险金额" value={formatCurrency(dashboard.kpis.riskAmount)} sub={`异常 ${formatNumber(dashboard.kpis.closeCount)} 笔`} icon={SlidersHorizontal} tone="red" />
        <KpiCard title="完成订单客单价" value={formatCurrency(dashboard.kpis.averageOrder)} sub={`曝光 ${formatNumber(dashboard.kpis.exposure)} · 点击 ${formatNumber(dashboard.kpis.clicks)}`} icon={FileSpreadsheet} tone="violet" />
      </section>

      <section className="story-grid">
        <ChartPanel title="商品成交贡献排行" note="按订单明细聚合成交与关闭风险" className="wide">
          <ResponsiveContainer {...RESPONSIVE_CHART_PROPS}>
            <BarChart data={dashboard.productBars.slice(0, 8)} margin={{ top: 16, right: 12, left: 10, bottom: 8 }}>
              <CartesianGrid stroke="#2b3437" vertical={false} />
              <XAxis dataKey="label" stroke="#91a0a5" interval={0} height={48} tickMargin={8} tick={<ProductAxisTick />} />
              <YAxis stroke="#91a0a5" tick={{ fontSize: 11 }} />
              <Tooltip content={<MiniTooltip />} cursor={{ fill: 'rgba(255,255,255,.04)' }} />
              <Legend />
              <Bar name="成交金额" dataKey="completed" fill="#10d2a0" radius={[5, 5, 0, 0]} />
              <Bar name="关闭/退款金额" dataKey="risk" fill="#7fb0d2" radius={[5, 5, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartPanel>

        <ChartPanel title="订单状态占比" note="识别完成与流失订单结构">
          <ResponsiveContainer {...RESPONSIVE_CHART_PROPS}>
            <PieChart>
              <Pie
                data={dashboard.statusPie}
                dataKey="value"
                nameKey="name"
                innerRadius={58}
                outerRadius={94}
                paddingAngle={0}
                stroke="#1b2022"
                strokeWidth={2}
                animationBegin={0}
                animationDuration={900}
                animationEasing="ease-in-out"
              >
                {dashboard.statusPie.map((entry) => <Cell key={entry.name} fill={entry.color} />)}
              </Pie>
              <Tooltip content={<MiniTooltip />} />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartPanel>
      </section>

      <section className="story-grid lower">
        <ChartPanel title="月度金额趋势" note="总金额、完成金额与风险金额对比" className="wide">
          <ResponsiveContainer {...RESPONSIVE_CHART_PROPS}>
            <ComposedChart data={dashboard.monthly} margin={{ top: 18, right: 22, left: 6, bottom: 8 }}>
              <defs>
                <linearGradient id="trendFill" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor="#7fb0d2" stopOpacity=".45" />
                  <stop offset="100%" stopColor="#7fb0d2" stopOpacity="0" />
                </linearGradient>
              </defs>
              <CartesianGrid stroke="#2b3437" vertical={false} />
              <XAxis dataKey="month" stroke="#91a0a5" tick={{ fontSize: 11 }} />
              <YAxis stroke="#91a0a5" tick={{ fontSize: 11 }} />
              <Tooltip content={<MiniTooltip />} />
              <Legend />
              <Area type="monotone" name="总金额" dataKey="amount" stroke="#7fb0d2" fill="url(#trendFill)" strokeWidth={2} />
              <Line type="monotone" name="完成金额" dataKey="completed" stroke="#10d2a0" strokeWidth={3} dot={{ r: 3 }} />
              <Line type="monotone" name="风险金额" dataKey="risk" stroke="#b8cdf4" strokeWidth={2} dot={{ r: 3 }} />
            </ComposedChart>
          </ResponsiveContainer>
        </ChartPanel>

        <section className="panel product-list">
          <div className="panel-heading">
            <div>
              <h2>商品销售/退款结构</h2>
              <p>结合商品价格总览与订单明细</p>
            </div>
            <CalendarClock size={17} />
          </div>
          <div className="rank-list">
            {dashboard.productMix.map((item, index) => (
              <div className="rank-row" key={item.product}>
                <div className="rank-meta">
                  <span className="rank-index">{String(index + 1).padStart(2, '0')}</span>
                  <div>
                    <b title={item.product}>{productAlias(item.product)}</b>
                    <small>销量 {formatNumber(item.soldQty)} · 退款 {formatNumber(item.refundQty)}</small>
                    {item.giftNote ? <em>{item.giftNote}</em> : null}
                  </div>
                </div>
                <div className="rank-bar">
                  <i style={{ width: `${Math.max(8, (item.amount / maxProductAmount) * 100)}%` }} />
                </div>
                <strong>{formatCurrency(item.amount)}</strong>
              </div>
            ))}
          </div>
        </section>
      </section>

      <section className="bottom-grid">
        <ChartPanel title="商品订单量分布" note="不同颜色代表不同商品">
          <ResponsiveContainer {...RESPONSIVE_CHART_PROPS}>
            <BarChart data={dashboard.productBars.slice(0, 8)} margin={{ top: 18, right: 18, left: 4, bottom: 8 }}>
              <CartesianGrid stroke="#2b3437" />
              <XAxis
                dataKey="label"
                stroke="#91a0a5"
                interval={0}
                height={48}
                tickMargin={10}
                tick={<ProductAxisTick />}
              />
              <YAxis stroke="#91a0a5" tick={{ fontSize: 11 }} />
              <Tooltip content={<MiniTooltip />} />
              <Bar name="订单量" dataKey="orders" radius={[6, 6, 0, 0]}>
                {dashboard.productBars.slice(0, 8).map(item => <Cell key={item.product} fill={item.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </ChartPanel>

        <section className="panel operations">
          <div className="panel-heading">
            <div>
              <h2>运营待办状态</h2>
              <p>读取主页数据中的风险与履约字段</p>
            </div>
            <SlidersHorizontal size={17} />
          </div>
          <div className="ops-grid">
            {['待支付', '待发货', '异常包裹', '待处理售后', '服务工单', '待整改风险点', '待处理违规'].map(item => (
              <div className="ops-cell" key={item}>
                <span>{item}</span>
                <b>{formatNumber(normalizeCount(dashboard.homeMetrics?.[item]))}</b>
              </div>
            ))}
          </div>
        </section>
      </section>
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
