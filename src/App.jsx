import { useState, useRef, useEffect } from 'react';
import PptxGenJS from 'pptxgenjs';
import DocToPptModal from './DocToPptModal';
import TemplateModal from './TemplateModal';
import { renderRichText, toPptxRichText, stripMarkdown, DEFAULT_TABLE_DATA } from './reportUtils';
import { CHART_TYPES, DEFAULT_CHART_DATA, isMultiSeries, ChartPreview, drawChartOnCanvas, addChartToPptxSlide } from './ChartEngine';

const STORAGE_KEY = 'ppt_studio_autosave_v1';

const SAFE_FONTS = [
  { name: '思源黑体（免费商用 · 替代微软雅黑首选）', value: 'Noto Sans SC' },
  { name: '思源宋体（免费商用 · 政务/高端汇报）', value: 'Noto Serif SC' },
  { name: '站酷小薇体（免费商用 · 品牌/文艺）', value: 'ZCOOL XiaoWei' },
  { name: '黑体（论文答辩常用 · 系统字体）', value: 'SimHei' },
  { name: '微软雅黑（论文答辩常用 · 系统字体）', value: 'Microsoft YaHei' },
  { name: '宋体（文科/法学答辩 · 系统字体）', value: 'SimSun' },
];

const THEMES = [
  { id: 'redblue', name: '经典红蓝汇报', isDark: false, isSplit: true, coverBgColor: '#0A2540', coverBgEnd: '#051526', bg: '#FFFFFF', bgEnd: '#F8FAFC', cardBg: '#F8FAFC', primary: '#0A3254', accent: '#C41E3A', text: '#0F172A', subText: '#475569' },
  { id: 'tech', name: '大厂科技蓝', isDark: true, bg: '#0B192C', bgEnd: '#060E1A', cardBg: '#132642', primary: '#38BDF8', accent: '#F59E0B', text: '#F8FAFC', subText: '#94A3B8' },
  { id: 'gov', name: '政务党建红', isDark: false, bg: '#FFFDF9', bgEnd: '#F8EFE6', cardBg: '#FFFFFF', primary: '#C41E3A', accent: '#D4AF37', text: '#1F2937', subText: '#6B7280' },
  { id: 'gold', name: '高端黑金风', isDark: true, bg: '#161616', bgEnd: '#0A0A0A', cardBg: '#222222', primary: '#D4AF37', accent: '#FDE68A', text: '#FFFFFF', subText: '#A3A3A3' },
  { id: 'minimal', name: '极简商务白', isDark: false, bg: '#F8FAFC', bgEnd: '#E2E8F0', cardBg: '#FFFFFF', primary: '#1D4ED8', accent: '#0EA5E9', text: '#0F172A', subText: '#64748B' },
];

const LAYOUTS = [
  { id: 'cover', name: '封面页' },
  { id: 'toc', name: '目录页（4章节）' },
  { id: 'section', name: '章节过渡页' },
  { id: 'dashboard', name: '📈 经营数据仪表盘（顶部KPI+多系列图）' },
  { id: 'table', name: '📋 商务数据表格页（汇报核心）' },
  { id: 'fourCol', name: '🗂️ 四栏/四季度汇报页（高密）' },
  { id: 'chart', name: '📊 多系列数据图表页（柱/折/环/饼）' },
  { id: 'compare', name: '左右对比页（痛点 vs 方案）' },
  { id: 'imageText', name: '左图右文页（产品/案例）' },
  { id: 'cards', name: '三栏并列卡片（带图标）' },
  { id: 'timeline', name: '时间轴/流程' },
  { id: 'metrics', name: '核心数据指标' },
  { id: 'ending', name: '致谢结尾页' },
];

const DEFAULT_KPIS = [
  { label: '全年签约总额', value: '1023亿', trend: '同比 +30.0%' },
  { label: '创造净利润', value: '141.4亿', trend: '同比 +37.4%' },
  { label: '滚动回款率', value: '96.3%', trend: '提升 +2.5pt' },
];

const ICONS = [
  { id: 'trend', name: '增长', paths: '<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/>' },
  { id: 'rocket', name: '启动', paths: '<path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.84.7-2.13-.09-2.91a2.18 2.18 0 0 0-2.91-.09z"/><path d="m12 15-3-3a22 22 0 0 1 2-3.95A12.88 12.88 0 0 1 22 2c0 2.72-.78 7.5-6 11a22.35 22.35 0 0 1-4 2z"/>' },
  { id: 'shield', name: '合规', paths: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/>' },
  { id: 'target', name: '精准', paths: '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>' },
  { id: 'zap', name: '效率', paths: '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>' },
  { id: 'award', name: '口碑', paths: '<circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/>' },
  { id: 'users', name: '团队', paths: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>' },
  { id: 'bulb', name: '创意', paths: '<path d="M9 18h6"/><path d="M10 22h4"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/>' },
  { id: 'layers', name: '架构', paths: '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>' },
  { id: 'clock', name: '周期', paths: '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>' },
  { id: 'chart', name: '数据', paths: '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>' },
  { id: 'globe', name: '全球', paths: '<circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>' },
];

const SENSITIVE_KEYWORDS = ['地图', '全国', '省市', '版图', '国旗', '国徽', '领土', '区域分布'];
const hex = (color) => (color || '#000000').replace('#', '');

const hexToRgba = (hexColor, alpha) => {
  const c = hex(hexColor);
  const r = parseInt(c.slice(0, 2), 16) || 0;
  const g = parseInt(c.slice(2, 4), 16) || 0;
  const b = parseInt(c.slice(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

const renderIconToDataURL = (iconId, color) =>
  new Promise((resolve) => {
    const icon = ICONS.find((ic) => ic.id === iconId) || ICONS[0];
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${icon.paths}</svg>`;
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas');
      c.width = 96;
      c.height = 96;
      c.getContext('2d').drawImage(img, 0, 0, 96, 96);
      resolve(c.toDataURL('image/png'));
    };
    img.onerror = () => resolve(null);
    img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  });

const generateTexturedBgDataURL = (theme, isDarkCover = false) => {
  const c = document.createElement('canvas');
  c.width = 1280;
  c.height = 720;
  const ctx = c.getContext('2d');
  const bgStart = isDarkCover ? theme.coverBgColor || theme.primary || '#0A2540' : theme.bg;
  const bgStop = isDarkCover ? theme.coverBgEnd || '#051526' : theme.bgEnd || theme.bg;
  const grad = ctx.createLinearGradient(0, 0, 1280, 720);
  grad.addColorStop(0, bgStart);
  grad.addColorStop(1, bgStop);
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, 1280, 720);

  const glow1 = ctx.createRadialGradient(1120, 90, 20, 1120, 90, 480);
  glow1.addColorStop(0, hexToRgba(isDarkCover ? theme.accent : theme.primary, isDarkCover || theme.isDark ? 0.16 : 0.08));
  glow1.addColorStop(1, hexToRgba(theme.primary, 0));
  ctx.fillStyle = glow1;
  ctx.fillRect(0, 0, 1280, 720);
  return c.toDataURL('image/jpeg', 0.92);
};

function IconSvg({ id, color, size = 22 }) {
  const icon = ICONS.find((ic) => ic.id === id) || ICONS[0];
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" dangerouslySetInnerHTML={{ __html: icon.paths }} />
  );
}

const INITIAL_SLIDES = [
  {
    id: 1,
    layout: 'cover',
    title: '2026 年度经营业绩与工作汇报',
    subtitle: '抗住压力，实现高质量增长 · 汇报人：您的名字',
    image: '',
    fontScale: 1,
    reverseLayout: false,
    notes: '各位领导、同事大家好，下面由我为大家汇报2026年度经营业绩与重点工作。',
    kpis: DEFAULT_KPIS,
    chartType: 'groupedBar',
    chartSeriesName: '2025年',
    chartSeriesName2: '2024年',
    chartUnit: '单位：亿元',
    chartData: DEFAULT_CHART_DATA,
    tableData: DEFAULT_TABLE_DATA,
    items: [
      { head: '要点一', body: '简明扼要写出核心内容。', icon: 'target', badge: '' },
      { head: '要点二', body: '简明扼要写出核心内容。', icon: 'zap', badge: '' },
      { head: '要点三', body: '简明扼要写出核心内容。', icon: 'award', badge: '' },
      { head: '要点四', body: '简明扼要写出核心内容。', icon: 'shield', badge: '' },
    ],
  },
  {
    id: 2,
    layout: 'dashboard',
    title: '经营数据仪表盘：核心指标连续四季度双位数增长',
    subtitle: '顶部 KPI 总览 + 下方双系列对比柱状图 + 右侧核心归因',
    image: '',
    fontScale: 1,
    reverseLayout: false,
    notes: '',
    kpis: DEFAULT_KPIS,
    chartType: 'groupedBar',
    chartSeriesName: '2025年',
    chartSeriesName2: '2024年',
    chartUnit: '单位：亿元',
    chartData: [
      { label: '第一季度', value: 210, value2: 165 },
      { label: '第二季度', value: 255, value2: 190 },
      { label: '第三季度', value: 268, value2: 205 },
      { label: '第四季度', value: 290, value2: 227 },
    ],
    tableData: DEFAULT_TABLE_DATA,
    items: [
      { head: '增长引擎：Q4 冲刺创新高', body: '第四季度单季完成 **290亿**，较去年同期大增 **+27.8%**，稳居全年峰值。', icon: 'trend', badge: '' },
      { head: '盈利质量：利润增速跑赢营收', body: '高毛利业务占比提升至 **42%**，推动全年创造净利润达 **141.4亿**。', icon: 'award', badge: '' },
      { head: '', body: '', icon: 'shield', badge: '' },
      { head: '', body: '', icon: 'zap', badge: '' },
    ],
  },
  {
    id: 3,
    layout: 'table',
    title: '整体经营数据回顾：2025年核心指标全面超额达成',
    subtitle: '支持 Excel 一键粘贴导入表格，导出 .pptx 后 100% 原生可编辑',
    image: '',
    fontScale: 1,
    reverseLayout: false,
    notes: '',
    kpis: DEFAULT_KPIS,
    chartType: 'line',
    chartSeriesName: '2025年',
    chartSeriesName2: '2024年',
    chartUnit: '单位：亿元',
    chartData: DEFAULT_CHART_DATA,
    tableData: DEFAULT_TABLE_DATA,
    items: [
      { head: '核心亮点一', body: '全年签约突破 **1023亿**，同比大增 **+30%**，回款率稳居 **96.3%**。', icon: 'trend', badge: '' },
      { head: '核心亮点二', body: '结算净利达 **111.4亿**（**+75.7%**），降本增效与高毛利项目开始释放利润。', icon: 'award', badge: '' },
      { head: '', body: '', icon: 'shield', badge: '' },
      { head: '', body: '', icon: 'zap', badge: '' },
    ],
  },
  {
    id: 4,
    layout: 'fourCol',
    title: '目标达成回顾：四大节点营销屡创新高，稳固全年业绩',
    subtitle: '大厂工作汇报高频四栏结构（顶部深蓝/正红标题块 + 多行要点 + 底部数据胶囊）',
    image: '',
    fontScale: 1,
    reverseLayout: false,
    notes: '',
    kpis: DEFAULT_KPIS,
    chartType: 'doughnut',
    chartSeriesName: '渠道占比',
    chartSeriesName2: '',
    chartUnit: '单位：万元',
    chartData: DEFAULT_CHART_DATA,
    tableData: DEFAULT_TABLE_DATA,
    items: [
      { head: 'Q1 新春大促', body: '• 聚焦品类专场与会员专享\n• 会员消费占比达 **65%**\n• 客户满意度提升至 **4.8分**', badge: 'GMV 3.25亿', icon: 'target' },
      { head: 'Q2 618大促', body: '• 明星推荐+跨界新品首发\n• 新品销量占比达 **40%**\n• 社媒总曝光超 **2亿次**', badge: 'GMV 3.85亿', icon: 'rocket' },
      { head: 'Q3 双11预热', body: '• 全品类直播带货+限时秒杀\n• 直播间销售额破 **1.5亿**\n• 新增用户转化率 **30%**', badge: 'GMV 3.20亿', icon: 'zap' },
      { head: 'Q4 年终清仓', body: '• 限量礼盒与老客复购激励\n• 库存周转率提升 **40%**\n• 获客成本降低 **20%**', badge: 'GMV 2.10亿', icon: 'award' },
    ],
  },
];

const loadSavedProject = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (parsed && Array.isArray(parsed.slides) && parsed.slides.length > 0) return parsed;
  } catch {
    return null;
  }
  return null;
};

const loadImage = (src) =>
  new Promise((resolve) => {
    if (!src) return resolve(null);
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => resolve(null);
    img.src = src;
  });

export default function App() {
  const saved = useRef(loadSavedProject()).current;

  const [projectName, setProjectName] = useState(saved?.projectName || '年度经营汇报方案');
  const [slides, setSlides] = useState(saved?.slides || INITIAL_SLIDES);
  const [past, setPast] = useState([]);
  const [future, setFuture] = useState([]);
  const [currentIdx, setCurrentIdx] = useState(0);

  const [themeId, setThemeId] = useState(saved?.themeId || 'redblue');
  const [customPrimary, setCustomPrimary] = useState(saved?.customPrimary || '');
  const [brandLogo, setBrandLogo] = useState(saved?.brandLogo || '');
  const [fontFamily, setFontFamily] = useState(saved?.fontFamily || 'Noto Sans SC');
  const [densityMode, setDensityMode] = useState(saved?.densityMode || 'report');
  const [showSafeGuide, setShowSafeGuide] = useState(true);
  const [showFooter, setShowFooter] = useState(saved?.showFooter ?? true);
  const [headerTag, setHeaderTag] = useState(saved?.headerTag || '');
  const [sourceNote, setSourceNote] = useState(saved?.sourceNote || '');
  const [watermarkText, setWatermarkText] = useState(saved?.watermarkText || '初稿预览 · 付尾款后去水印');

  const [toastMsg, setToastMsg] = useState('');
  const [saveStatus, setSaveStatus] = useState(saved ? '已恢复上次存档' : '自动存档已开启');
  const [isExportingLong, setIsExportingLong] = useState(false);
  const [showDocModal, setShowDocModal] = useState(false);
  const [coverBg, setCoverBg] = useState(saved?.coverBg || '');
  const [contentBg, setContentBg] = useState(saved?.contentBg || '');
  const [showTemplateModal, setShowTemplateModal] = useState(false);
  const [showAuditModal, setShowAuditModal] = useState(false);
  const [showExcelPaste, setShowExcelPaste] = useState(false);
  const [excelRawText, setExcelRawText] = useState('');

  const logoInputRef = useRef(null);
  const slideImgInputRef = useRef(null);
  const jsonInputRef = useRef(null);

  const baseTheme = THEMES.find((t) => t.id === themeId) || THEMES[0];
  const theme = { ...baseTheme, primary: customPrimary || baseTheme.primary };
  const currentSlide = slides[currentIdx] || slides[0];
  const currentChartData = currentSlide.chartData && currentSlide.chartData.length > 0 ? currentSlide.chartData : DEFAULT_CHART_DATA;
  const currentTableData = currentSlide.tableData && currentSlide.tableData.headers ? currentSlide.tableData : DEFAULT_TABLE_DATA;
  const currentKpis = currentSlide.kpis && currentSlide.kpis.length >= 3 ? currentSlide.kpis : DEFAULT_KPIS;
  const currentChartType = currentSlide.chartType || 'groupedBar';
  const currentFontScale = currentSlide.fontScale || 1;
  const isReversed = Boolean(currentSlide.reverseLayout);
  const showChartEditor = currentSlide.layout === 'chart' || currentSlide.layout === 'dashboard';
  const maxCharLimit = densityMode === 'report' ? 380 : 120;

  const isContentSlide = !['cover', 'section', 'ending'].includes(currentSlide.layout);
  const isDarkCover = !isContentSlide && (theme.isSplit || theme.id === 'minimal');
  const activeText = isDarkCover ? '#FFFFFF' : theme.text;
  const activeSubText = isDarkCover ? '#CBD5E1' : theme.subText;

  const cardShadow = theme.isDark
    ? '0 10px 28px rgba(0, 0, 0, 0.42)'
    : '0 6px 20px rgba(15, 23, 42, 0.06), 0 1px 4px rgba(15, 23, 42, 0.04)';

  const canvasBackgroundCss = isDarkCover
    ? `radial-gradient(circle at 86% 18%, ${hexToRgba(theme.accent, 0.22)}, transparent 45%), linear-gradient(135deg, ${theme.coverBgColor || theme.primary} 0%, ${theme.coverBgEnd || '#051526'} 100%)`
    : theme.isSplit
      ? '#FFFFFF'
      : `radial-gradient(circle at 88% 12%, ${hexToRgba(theme.primary, theme.isDark ? 0.16 : 0.1)}, transparent 44%), linear-gradient(135deg, ${theme.bg} 0%, ${theme.bgEnd} 100%)`;

  // ================= 撤销 / 重做历史栈 =================
  const commitSlides = (updater) => {
    setSlides((prev) => {
      const next = typeof updater === 'function' ? updater(prev) : updater;
      setPast((p) => [...p.slice(-19), prev]);
      setFuture([]);
      return next;
    });
  };

  const handleUndo = () => {
    if (past.length === 0) return;
    const previous = past[past.length - 1];
    setPast((p) => p.slice(0, -1));
    setFuture((f) => [slides, ...f.slice(0, 19)]);
    setSlides(previous);
    setCurrentIdx((idx) => Math.min(idx, previous.length - 1));
    showToast('↩️ 已撤销上一步操作');
  };

  const handleRedo = () => {
    if (future.length === 0) return;
    const next = future[0];
    setFuture((f) => f.slice(1));
    setPast((p) => [...p.slice(-19), slides]);
    setSlides(next);
    setCurrentIdx((idx) => Math.min(idx, next.length - 1));
    showToast('↪️ 已重做');
  };

  useEffect(() => {
    try {
      const payload = { version: 6, projectName, themeId, customPrimary, brandLogo, fontFamily, densityMode, showFooter, headerTag, sourceNote, coverBg, contentBg, watermarkText, slides };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      setSaveStatus(`已自动保存 ${new Date().toTimeString().slice(0, 8)}`);
    } catch {
      setSaveStatus('⚠️ 图片过大，请点「导出订单」手动备份');
    }
  }, [projectName, slides, themeId, customPrimary, brandLogo, fontFamily, densityMode, showFooter, headerTag, sourceNote, coverBg, contentBg, watermarkText]);

  const showToast = (msg) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg((prev) => (prev === msg ? '' : prev)), 4500);
  };

  const handleExportProjectJSON = () => {
    const payload = { version: 6, exportedAt: new Date().toISOString(), projectName, themeId, customPrimary, brandLogo, fontFamily, densityMode, showFooter, headerTag, sourceNote, coverBg, contentBg, watermarkText, slides };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${projectName || '未命名项目'}-接单存档.json`;
    link.click();
    showToast(`💾 已导出订单存档「${projectName}-接单存档.json」。`);
  };

  const handleImportProjectJSON = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!Array.isArray(data.slides) || data.slides.length === 0) throw new Error('无效文件');
        setProjectName(data.projectName || file.name.replace(/\.json$/i, ''));
        commitSlides(data.slides);
        setThemeId(data.themeId || 'redblue');
        setCustomPrimary(data.customPrimary || '');
        setBrandLogo(data.brandLogo || '');
        setFontFamily(data.fontFamily || 'Noto Sans SC');
        if (data.densityMode) setDensityMode(data.densityMode);
        if (typeof data.showFooter === 'boolean') setShowFooter(data.showFooter);
        setHeaderTag(data.headerTag || '');
        setSourceNote(data.sourceNote || '');
        setCoverBg(data.coverBg || '');
        setContentBg(data.contentBg || '');
        setWatermarkText(data.watermarkText || '初稿预览 · 付尾款后去水印');
        setCurrentIdx(0);
        showToast(`✅ 已载入订单「${data.projectName || file.name}」（共 ${data.slides.length} 页）`);
      } catch {
        alert('导入失败：请选择本软件导出的 .json 订单存档文件。');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const handleResetProject = () => {
    if (!window.confirm('确定要新建空白汇报订单吗？\n（当前未导出 .json 的内容将被重置为默认模板）')) return;
    setProjectName('年度经营汇报方案');
    commitSlides(INITIAL_SLIDES);
    setThemeId('redblue');
    setCustomPrimary('');
    setBrandLogo('');
    setHeaderTag('');
    setSourceNote('');
    setCoverBg('');
    setContentBg('');
    setFontFamily('Noto Sans SC');
    setCurrentIdx(0);
    showToast('✨ 已重置为全新红蓝汇报模板。');
  };

  const updateSlide = (patch) => {
    commitSlides((prev) => prev.map((s, idx) => (idx === currentIdx ? { ...s, ...patch } : s)));
  };

  const updateItem = (itemIdx, field, val) => {
    const baseItems = [...(currentSlide.items || [])];
    while (baseItems.length <= itemIdx) baseItems.push({ head: '', body: '', icon: ICONS[itemIdx % ICONS.length].id, badge: '' });
    const newItems = baseItems.map((item, i) => (i === itemIdx ? { ...item, [field]: val } : item));
    updateSlide({ items: newItems });
  };

  const updateKpi = (kpiIdx, field, val) => {
    const nextKpis = currentKpis.map((k, i) => (i === kpiIdx ? { ...k, [field]: val } : k));
    updateSlide({ kpis: nextKpis });
  };

  // ================= Excel / WPS 一键粘贴解析 =================
  const handleApplyExcelPaste = (mode) => {
    const lines = excelRawText.trim().split(/\r?\n/).filter((l) => l.trim());
    if (lines.length < 2) return alert('请至少粘贴 2 行数据（直接从 Excel 或 WPS 表格复制区域即可）');
    const matrix = lines.map((line) => line.split(/\t|,/).map((cell) => cell.trim()));

    if (mode === 'table') {
      const headers = matrix[0].slice(0, 6);
      const rows = matrix.slice(1, 8).map((r) => headers.map((_, ci) => r[ci] ?? '-'));
      updateSlide({ tableData: { headers, rows } });
      showToast(`📋 已从 Excel 导入 ${rows.length} 行 × ${headers.length} 列表格！`);
    } else {
      const hasHeader = isNaN(Number(matrix[0][1]));
      const dataLines = hasHeader ? matrix.slice(1, 7) : matrix.slice(0, 6);
      const chartData = dataLines.map((r) => ({
        label: r[0] || '项',
        value: Number(String(r[1] || '0').replace(/[^\d.-]/g, '')) || 0,
        value2: r[2] !== undefined ? Number(String(r[2]).replace(/[^\d.-]/g, '')) || 0 : 0,
      }));
      const patch = { chartData };
      if (hasHeader && matrix[0][1]) patch.chartSeriesName = matrix[0][1];
      if (hasHeader && matrix[0][2]) patch.chartSeriesName2 = matrix[0][2];
      updateSlide(patch);
      showToast(`📊 已从 Excel 导入 ${chartData.length} 组图表数据！`);
    }
    setExcelRawText('');
    setShowExcelPaste(false);
  };

  // ================= 表格与图表编辑 =================
  const updateTableHeader = (colIdx, val) => {
    const headers = currentTableData.headers.map((h, i) => (i === colIdx ? val : h));
    updateSlide({ tableData: { ...currentTableData, headers } });
  };

  const updateTableCell = (rowIdx, colIdx, val) => {
    const rows = currentTableData.rows.map((r, ri) => (ri === rowIdx ? r.map((c, ci) => (ci === colIdx ? val : c)) : r));
    updateSlide({ tableData: { ...currentTableData, rows } });
  };

  const addTableRow = () => {
    if (currentTableData.rows.length >= 7) return alert('建议单页表格最多 7 行数据');
    updateSlide({ tableData: { ...currentTableData, rows: [...currentTableData.rows, currentTableData.headers.map((_, i) => (i === 0 ? '新增指标' : '0'))] } });
  };

  const removeTableRow = (rowIdx) => {
    if (currentTableData.rows.length <= 1) return alert('表格至少保留 1 行数据');
    updateSlide({ tableData: { ...currentTableData, rows: currentTableData.rows.filter((_, i) => i !== rowIdx) } });
  };

  const addTableCol = () => {
    if (currentTableData.headers.length >= 6) return alert('建议最多 6 列');
    updateSlide({ tableData: { headers: [...currentTableData.headers, `列${currentTableData.headers.length + 1}`], rows: currentTableData.rows.map((r) => [...r, '-']) } });
  };

  const removeTableCol = () => {
    if (currentTableData.headers.length <= 2) return alert('表格至少保留 2 列');
    updateSlide({ tableData: { headers: currentTableData.headers.slice(0, -1), rows: currentTableData.rows.map((r) => r.slice(0, -1)) } });
  };

  const updateChartRow = (rowIdx, field, val) => {
    const nextData = currentChartData.map((row, i) => (i === rowIdx ? { ...row, [field]: field === 'label' ? val : Number(val) } : row));
    updateSlide({ chartData: nextData });
  };

  const addChartRow = () => {
    if (currentChartData.length >= 6) return alert('单图建议最多 6 组数据');
    updateSlide({ chartData: [...currentChartData, { label: `项${currentChartData.length + 1}`, value: 150, value2: 120 }] });
  };

  const removeChartRow = (rowIdx) => {
    if (currentChartData.length <= 2) return alert('图表至少保留 2 组数据');
    updateSlide({ chartData: currentChartData.filter((_, i) => i !== rowIdx) });
  };

  const addSlide = (layout = 'dashboard') => {
    const newSlide = {
      id: Date.now(),
      layout,
      title: layout === 'dashboard' ? '经营数据仪表盘：核心指标稳步增长' : layout === 'table' ? '经营数据汇总表' : '点击输入本页核心结论标题',
      subtitle: layout === 'section' ? 'PART 01' : '支持使用 **双星号** 自动加粗标红核心数据',
      image: '',
      fontScale: 1,
      reverseLayout: false,
      notes: '',
      kpis: DEFAULT_KPIS.map((k) => ({ ...k })),
      chartType: 'groupedBar',
      chartSeriesName: '2025年',
      chartSeriesName2: '2024年',
      chartUnit: '单位：万元',
      chartData: DEFAULT_CHART_DATA.map((d) => ({ ...d })),
      tableData: { headers: [...DEFAULT_TABLE_DATA.headers], rows: DEFAULT_TABLE_DATA.rows.map((r) => [...r]) },
      items: [
        { head: '核心结论一', body: '关键业务指标同比提升 **+32%**，增长动能强劲。', icon: 'trend', badge: '达成率 105%' },
        { head: '核心结论二', body: '降本增效显著，整体毛利率较年初提升 **+4.2pt**。', icon: 'award', badge: '达成率 112%' },
        { head: '要点三', body: '• 核心举措与成果', icon: 'zap', badge: '达成率 108%' },
        { head: '要点四', body: '• 全年目标超额完成', icon: 'shield', badge: '达成率 120%' },
      ],
    };
    commitSlides((prev) => [...prev, newSlide]);
    setCurrentIdx(slides.length);
  };

  const moveSlide = (idx, direction) => {
    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= slides.length) return;
    commitSlides((prev) => {
      const next = [...prev];
      const [moved] = next.splice(idx, 1);
      next.splice(targetIdx, 0, moved);
      return next;
    });
    if (currentIdx === idx) setCurrentIdx(targetIdx);
    else if (currentIdx === targetIdx) setCurrentIdx(idx);
  };

  const duplicateSlide = (idx) => {
    const source = slides[idx];
    const cloned = {
      ...source,
      id: Date.now(),
      title: `${source.title} (副本)`,
      kpis: (source.kpis || DEFAULT_KPIS).map((k) => ({ ...k })),
      items: (source.items || []).map((item) => ({ ...item })),
      chartData: (source.chartData || DEFAULT_CHART_DATA).map((d) => ({ ...d })),
      tableData: source.tableData ? { headers: [...source.tableData.headers], rows: source.tableData.rows.map((r) => [...r]) } : DEFAULT_TABLE_DATA,
    };
    commitSlides((prev) => {
      const next = [...prev];
      next.splice(idx + 1, 0, cloned);
      return next;
    });
    setCurrentIdx(idx + 1);
    showToast(`❐ 已复制第 ${idx + 1} 页。`);
  };

  const deleteSlide = (idx) => {
    if (slides.length <= 1) return alert('至少需要保留 1 页幻灯片');
    commitSlides((prev) => prev.filter((_, i) => i !== idx));
    setCurrentIdx(Math.max(0, idx - 1));
  };

  const handleLogoUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      setBrandLogo(ev.target.result);
      showToast('🛡️ 合规提醒：已上传客户 Logo，请确认已获官方授权。');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const handleSlideImageUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      updateSlide({ image: ev.target.result });
      showToast('🛡️ 图片合规提醒：请确认图片无版权风险。');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const getVisibleItemCount = (layout) => {
    if (layout === 'cover' || layout === 'ending' || layout === 'section') return 0;
    if (layout === 'compare' || layout === 'chart' || layout === 'table' || layout === 'dashboard') return 2;
    if (layout === 'toc' || layout === 'fourCol') return 4;
    return 3;
  };

  const supportsIcons = (layout) => ['cards', 'imageText'].includes(layout);
  const supportsMirror = (layout) => ['imageText', 'chart', 'dashboard'].includes(layout);

  const analyzeSlideCompliance = (slide) => {
    const warnings = [];
    const visibleCount = getVisibleItemCount(slide.layout);
    const activeItems = (slide.items || []).slice(0, visibleCount);
    const tableTexts = slide.layout === 'table' && slide.tableData ? [...slide.tableData.headers, ...slide.tableData.rows.flat()] : [];
    const kpiTexts = slide.layout === 'dashboard' && slide.kpis ? slide.kpis.flatMap((k) => [k.label, k.value, k.trend]) : [];

    const allText = stripMarkdown([slide.title || '', slide.subtitle || '', ...activeItems.flatMap((i) => [i.head || '', i.body || '', i.badge || '']), ...tableTexts, ...kpiTexts].join(' '));
    const charCount = allText.replace(/\s+/g, '').length;

    if (charCount > maxCharLimit) warnings.push({ level: 'warn', msg: `字数超出当前模式建议（${charCount} / ${maxCharLimit} 字）` });
    const hitSensitive = SENSITIVE_KEYWORDS.filter((kw) => allText.includes(kw));
    if (hitSensitive.length > 0) warnings.push({ level: 'danger', msg: `检测到敏感词「${hitSensitive.join('、')}」：严禁网图搜地图/国旗，必须使用自然资源部标准素材！` });
    if (slide.layout === 'imageText' && !slide.image) warnings.push({ level: 'info', msg: '当前为「左图右文」版式，尚未上传配图。' });
    if (allText.includes('点击输入') || allText.includes('您的名字')) warnings.push({ level: 'info', msg: '包含未替换的默认占位文字，交付前请修改。' });

    return { charCount, warnings };
  };

  const currentAudit = analyzeSlideCompliance(currentSlide);
  const fullDeckAudits = slides.map((s, idx) => ({ idx, title: s.title, ...analyzeSlideCompliance(s) }));
  const totalIssueCount = fullDeckAudits.reduce((sum, a) => sum + a.warnings.length, 0);

  // ================= 导出 1：零弹窗商业级可编辑 .pptx =================
  const handleExportPPTX = async () => {
    let dangerCount = 0;
    let warnCount = 0;
    fullDeckAudits.forEach((a) => a.warnings.forEach((w) => (w.level === 'danger' ? dangerCount++ : warnCount++)));

    if (dangerCount > 0 || warnCount > 0) {
      if (!window.confirm(`【导出前合规体检报告】\n检测到 ${dangerCount} 项红线风险，${warnCount} 项提醒。\n\n是否继续导出？`)) return;
    }

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    const bgDataUrl = generateTexturedBgDataURL(theme, false);
    const accentHex = hex(theme.accent);
    // PowerPoint 原生字体只取首个主字体名，确保 100% 识别
    const pptxFont = fontFamily.split(',')[0].replace(/['"]/g, '').trim() || 'Noto Sans SC';

    for (let sIdx = 0; sIdx < slides.length; sIdx++) {
      const s = slides[sIdx];
      const slide = pptx.addSlide();
      const isContent = !['cover', 'section', 'ending'].includes(s.layout);
      const isDark = !isContent && (theme.isSplit || theme.id === 'minimal');
      const customBg = isContent ? contentBg : coverBg;
      const fScale = s.fontScale || 1;
      const rev = Boolean(s.reverseLayout);

      if (customBg) {
        slide.background = { data: customBg };
      } else if (isDark) {
        const coverTheme = theme.coverBgColor ? theme : { ...theme, coverBgColor: theme.primary, coverBgEnd: '#0F172A' };
        slide.background = { data: generateTexturedBgDataURL(coverTheme, true) };
      } else if (theme.isSplit && isContent) {
        slide.background = { color: 'FFFFFF' };
      } else {
        slide.background = { data: bgDataUrl };
      }

      if (brandLogo) {
        slide.addImage({ data: brandLogo, x: 8.45, y: 0.25, w: 1.15, h: 0.48, sizing: { type: 'contain', w: 1.15, h: 0.48 } });
      } else if (headerTag && isContent) {
        slide.addText(headerTag, { x: 6.0, y: 0.24, w: 3.4, h: 0.3, fontSize: 9.5, fontFace: pptxFont, color: hex(theme.subText), align: 'right' });
      }

      if (s.notes && s.notes.trim()) {
        slide.addNotes(s.notes.trim());
      }

      const items = s.items || [];
      const sText = isDark ? 'FFFFFF' : hex(theme.text);
      const sSubText = isDark ? 'CBD5E1' : hex(theme.subText);

      if (showFooter && isContent) {
        slide.addShape(pptx.ShapeType.rect, { x: 0.6, y: 5.18, w: 8.8, h: 0.015, fill: { color: 'CBD5E1' } });
        slide.addText(sourceNote || `${projectName || '商业演示方案'}  ·  CONFIDENTIAL`, { x: 0.6, y: 5.24, w: 6.8, h: 0.28, fontSize: 9.5, fontFace: pptxFont, color: hex(theme.subText) });
        slide.addText(`${String(sIdx + 1).padStart(2, '0')} / ${String(slides.length).padStart(2, '0')}`, { x: 7.4, y: 5.24, w: 2.0, h: 0.28, fontSize: 10, bold: true, fontFace: pptxFont, color: hex(theme.primary), align: 'right' });
      }

      if (s.layout === 'cover') {
        slide.addText('ANNUAL REPORT / PRESENTATION', { x: 1.0, y: 1.25, w: 6.0, h: 0.35, fontSize: 12, bold: true, charSpacing: 4, color: isDark ? accentHex : hex(theme.primary) });
        slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.7, w: 0.12, h: 2.1, fill: { color: isDark ? accentHex : hex(theme.primary) } });
        slide.addText(s.title, { x: 1.0, y: 1.65, w: 8.2, h: 1.3, fontSize: Math.round(38 * fScale), bold: true, fontFace: pptxFont, color: sText, valign: 'middle', shrinkText: true });
        slide.addText(s.subtitle, { x: 1.0, y: 3.0, w: 8.2, h: 0.7, fontSize: Math.round(17 * fScale), fontFace: pptxFont, color: sSubText });
      } else if (s.layout === 'section') {
        slide.addText(s.subtitle || 'PART 01', { x: 0.8, y: 1.45, w: 8.4, h: 0.7, fontSize: 28, bold: true, charSpacing: 3, fontFace: pptxFont, color: isDark ? accentHex : hex(theme.primary), align: 'center' });
        slide.addShape(pptx.ShapeType.rect, { x: 4.3, y: 2.25, w: 1.4, h: 0.08, fill: { color: accentHex } });
        slide.addText(s.title, { x: 0.8, y: 2.5, w: 8.4, h: 1.2, fontSize: Math.round(40 * fScale), bold: true, fontFace: pptxFont, color: sText, align: 'center', valign: 'middle', shrinkText: true });
      } else if (s.layout === 'ending') {
        slide.addText('THANK YOU FOR WATCHING', { x: 0.8, y: 1.45, w: 8.4, h: 0.35, fontSize: 12, bold: true, charSpacing: 4, color: sSubText, align: 'center' });
        slide.addText(s.title || '感谢聆听', { x: 0.8, y: 1.9, w: 8.4, h: 1.1, fontSize: Math.round(42 * fScale), bold: true, fontFace: pptxFont, color: isDark ? 'FFFFFF' : hex(theme.primary), align: 'center', shrinkText: true });
        slide.addText(s.subtitle, { x: 0.8, y: 3.1, w: 8.4, h: 0.6, fontSize: 17, fontFace: pptxFont, color: sSubText, align: 'center' });
      } else {
        slide.addShape(pptx.ShapeType.rect, { x: 0.6, y: 0.42, w: 0.09, h: 0.44, fill: { color: accentHex } });
        slide.addText(s.title, { x: 0.78, y: 0.35, w: 7.4, h: 0.56, fontSize: Math.round(22 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.text), shrinkText: true });
        if (s.subtitle) {
          slide.addText(stripMarkdown(s.subtitle), { x: 0.78, y: 0.91, w: 8.2, h: 0.32, fontSize: 11.5, fontFace: pptxFont, color: hex(theme.subText), shrinkText: true });
        }

        if (s.layout === 'dashboard') {
          const kpis = s.kpis && s.kpis.length >= 3 ? s.kpis : DEFAULT_KPIS;
          kpis.slice(0, 3).forEach((k, i) => {
            const x = 0.6 + i * 3.0;
            const barHex = i === 1 ? accentHex : hex(theme.primary);
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.35, w: 2.8, h: 0.94, rectRadius: 0.06, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 1.35, w: 0.08, h: 0.94, fill: { color: barHex } });
            slide.addText(k.label, { x: x + 0.16, y: 1.41, w: 2.55, h: 0.26, fontSize: 10, fontFace: pptxFont, color: hex(theme.subText), shrinkText: true });
            slide.addText(k.value, { x: x + 0.16, y: 1.69, w: 1.35, h: 0.5, fontSize: 20, bold: true, fontFace: pptxFont, color: barHex, valign: 'middle', shrinkText: true });
            if (k.trend) {
              slide.addShape(pptx.ShapeType.roundRect, { x: x + 1.52, y: 1.8, w: 1.18, h: 0.3, rectRadius: 0.12, fill: { color: 'F1F5F9' } });
              slide.addText(k.trend, { x: x + 1.52, y: 1.8, w: 1.18, h: 0.3, fontSize: 8.5, bold: true, fontFace: pptxFont, color: barHex, align: 'center', valign: 'middle', shrinkText: true });
            }
          });

          const chartX = rev ? 3.85 : 0.6;
          const textX = rev ? 0.6 : 6.35;
          addChartToPptxSlide(pptx, slide, {
            x: chartX, y: 2.54, w: 5.55, h: 2.48,
            type: s.chartType || 'groupedBar', data: s.chartData, theme,
            series1: s.chartSeriesName, series2: s.chartSeriesName2, unit: s.chartUnit,
          });
          items.slice(0, 2).forEach((item, i) => {
            const y = 2.48 + i * 1.29;
            slide.addShape(pptx.ShapeType.roundRect, { x: textX, y, w: 3.05, h: 1.18, rectRadius: 0.06, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x: textX, y, w: 3.05, h: 0.05, fill: { color: i === 0 ? hex(theme.primary) : accentHex } });
            slide.addText(item.head, { x: textX + 0.14, y: y + 0.09, w: 2.77, h: 0.28, fontSize: Math.round(12 * fScale), bold: true, fontFace: pptxFont, color: i === 0 ? hex(theme.primary) : accentHex, shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(10.5 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: textX + 0.14, y: y + 0.39, w: 2.77, h: 0.72, valign: 'top', shrinkText: true });
          });
        } else if (s.layout === 'table') {
          const tData = s.tableData && s.tableData.headers ? s.tableData : DEFAULT_TABLE_DATA;
          const rowCount = tData.rows.length + 1;
          const dynamicRowH = Math.min(0.4, 2.3 / rowCount);
          const pptxRows = [
            tData.headers.map((h) => ({ text: h, options: { fill: { color: hex(theme.primary) }, color: 'FFFFFF', bold: true, fontSize: Math.round(11.5 * fScale), fontFace: pptxFont, align: 'center', valign: 'middle' } })),
            ...tData.rows.map((row, rIdx) =>
              row.map((cell, cIdx) => {
                const isHL = cIdx === row.length - 1 || /^\+/.test(String(cell).trim());
                return { text: String(cell), options: { fill: { color: rIdx % 2 === 1 ? 'F1F5F9' : 'FFFFFF' }, color: isHL ? accentHex : hex(theme.text), bold: cIdx === 0 || isHL, fontSize: Math.round(11 * fScale), fontFace: pptxFont, align: cIdx === 0 ? 'left' : 'center', valign: 'middle' } };
              })
            ),
          ];
          slide.addTable(pptxRows, { x: 0.6, y: 1.38, w: 8.8, rowH: dynamicRowH, border: { pt: 1, color: 'CBD5E1' } });
          items.slice(0, 2).forEach((item, i) => {
            const x = 0.6 + i * 4.55;
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 3.92, w: 4.25, h: 1.08, rectRadius: 0.06, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 3.92, w: 0.08, h: 1.08, fill: { color: i === 0 ? hex(theme.primary) : accentHex } });
            slide.addText(item.head, { x: x + 0.18, y: 3.99, w: 3.9, h: 0.28, fontSize: Math.round(12.5 * fScale), bold: true, fontFace: pptxFont, color: i === 0 ? hex(theme.primary) : accentHex, shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(10.5 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: x + 0.18, y: 4.3, w: 3.9, h: 0.64, valign: 'top', shrinkText: true });
          });
        } else if (s.layout === 'fourCol') {
          items.slice(0, 4).forEach((item, i) => {
            const x = 0.6 + i * 2.26;
            const headColor = i % 2 === 0 ? hex(theme.primary) : accentHex;
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.45, w: 2.08, h: 3.5, rectRadius: 0.06, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 1.45, w: 2.08, h: 0.46, fill: { color: headColor } });
            slide.addText(item.head, { x, y: 1.45, w: 2.08, h: 0.46, fontSize: Math.round(13 * fScale), bold: true, fontFace: pptxFont, color: 'FFFFFF', align: 'center', valign: 'middle', shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(10.5 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: x + 0.1, y: 2.02, w: 1.88, h: 2.26, valign: 'top', lineSpacingMultiple: 1.2, shrinkText: true });
            if (item.badge) {
              slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.16, y: 4.4, w: 1.76, h: 0.36, rectRadius: 0.15, fill: { color: headColor } });
              slide.addText(item.badge, { x: x + 0.16, y: 4.4, w: 1.76, h: 0.36, fontSize: 10.5, bold: true, fontFace: pptxFont, color: 'FFFFFF', align: 'center', valign: 'middle', shrinkText: true });
            }
          });
        } else if (s.layout === 'chart') {
          const chartX = rev ? 3.9 : 0.6;
          const textX = rev ? 0.6 : 6.35;
          addChartToPptxSlide(pptx, slide, {
            x: chartX, y: 1.52, w: 5.5, h: 3.4,
            type: s.chartType || 'groupedBar', data: s.chartData, theme,
            series1: s.chartSeriesName, series2: s.chartSeriesName2, unit: s.chartUnit,
          });
          items.slice(0, 2).forEach((item, i) => {
            const y = 1.48 + i * 1.74;
            slide.addShape(pptx.ShapeType.roundRect, { x: textX, y, w: 3.05, h: 1.58, rectRadius: 0.08, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x: textX, y, w: 3.05, h: 0.06, fill: { color: hex(theme.primary) } });
            slide.addText(item.head, { x: textX + 0.18, y: y + 0.16, w: 2.68, h: 0.38, fontSize: Math.round(13.5 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.primary), shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(11 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: textX + 0.18, y: y + 0.58, w: 2.68, h: 0.9, valign: 'top', shrinkText: true });
          });
        } else if (s.layout === 'toc') {
          items.slice(0, 4).forEach((item, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const x = 0.6 + col * 4.55;
            const y = 1.58 + row * 1.66;
            slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 4.25, h: 1.42, rectRadius: 0.08, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: y + 0.25, w: 0.08, h: 0.92, fill: { color: hex(theme.primary) } });
            slide.addText(`0${i + 1}`, { x: x + 0.24, y: y + 0.2, w: 0.8, h: 0.92, fontSize: 28, bold: true, fontFace: pptxFont, color: hex(theme.primary), valign: 'middle' });
            slide.addText(item.head, { x: x + 1.1, y: y + 0.22, w: 2.95, h: 0.44, fontSize: Math.round(15.5 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.text), shrinkText: true });
            slide.addText(stripMarkdown(item.body), { x: x + 1.1, y: y + 0.72, w: 2.95, h: 0.56, fontSize: Math.round(12 * fScale), fontFace: pptxFont, color: hex(theme.subText), shrinkText: true });
          });
        } else if (s.layout === 'imageText') {
          const imgX = rev ? 5.1 : 0.6;
          const txtX = rev ? 0.6 : 5.2;
          if (s.image) {
            slide.addImage({ data: s.image, x: imgX, y: 1.52, w: 4.3, h: 3.35, sizing: { type: 'contain', w: 4.3, h: 3.35 } });
          } else {
            slide.addShape(pptx.ShapeType.roundRect, { x: imgX, y: 1.52, w: 4.3, h: 3.35, rectRadius: 0.08, fill: { color: hex(theme.cardBg) }, line: { color: 'CBD5E1', width: 1 } });
            slide.addText('配图占位区', { x: imgX, y: 2.9, w: 4.3, h: 0.6, fontSize: 16, fontFace: pptxFont, color: hex(theme.subText), align: 'center' });
          }
          for (let i = 0; i < 3; i++) {
            const item = items[i] || { head: '', body: '', icon: ICONS[i].id };
            const y = 1.52 + i * 1.14;
            const iconPng = await renderIconToDataURL(item.icon || ICONS[i].id, theme.primary);
            if (iconPng) {
              slide.addShape(pptx.ShapeType.ellipse, { x: txtX, y: y + 0.06, w: 0.44, h: 0.44, fill: { color: 'E2E8F0' } });
              slide.addImage({ data: iconPng, x: txtX + 0.09, y: y + 0.15, w: 0.26, h: 0.26 });
            }
            slide.addText(item.head, { x: txtX + 0.58, y, w: 3.6, h: 0.34, fontSize: Math.round(14.5 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.primary), shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(11.5 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: txtX + 0.58, y: y + 0.36, w: 3.6, h: 0.72, valign: 'top', shrinkText: true });
          }
        } else if (s.layout === 'compare') {
          items.slice(0, 2).forEach((item, i) => {
            const x = 0.6 + i * 4.55;
            const barColor = i === 0 ? hex(theme.subText) : hex(theme.primary);
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.48, w: 4.25, h: 3.45, rectRadius: 0.08, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 1.48, w: 4.25, h: 0.07, fill: { color: barColor } });
            slide.addText(item.head, { x: x + 0.22, y: 1.68, w: 3.8, h: 0.44, fontSize: Math.round(16.5 * fScale), bold: true, fontFace: pptxFont, color: barColor, shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(10 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: x + 0.2, y: 2.15, w: 3.85, h: 2.72, valign: 'top', lineSpacingMultiple: 1.15 });
          });
        } else if (s.layout === 'cards') {
          for (let i = 0; i < 3; i++) {
            const item = items[i] || { head: '', body: '', icon: ICONS[i].id };
            const x = 0.6 + i * 3.05;
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.48, w: 2.75, h: 3.45, rectRadius: 0.08, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 1.48, w: 2.75, h: 0.07, fill: { color: hex(theme.primary) } });
            const iconPng = await renderIconToDataURL(item.icon || ICONS[i].id, theme.primary);
            if (iconPng) {
              slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.2, y: 1.7, w: 0.46, h: 0.46, rectRadius: 0.1, fill: { color: 'E2E8F0' } });
              slide.addImage({ data: iconPng, x: x + 0.28, y: 1.78, w: 0.3, h: 0.3 });
            }
            slide.addText(item.head, { x: x + 0.2, y: 2.28, w: 2.35, h: 0.4, fontSize: Math.round(15 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.text), shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(9.5 * fScale), fontFace: pptxFont, color: hex(theme.subText) }, accentHex), { x: x + 0.16, y: 2.66, w: 2.43, h: 2.22, valign: 'top', lineSpacingMultiple: 1.15 });
          }
        } else if (s.layout === 'timeline') {
          slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.55, w: 8.4, h: 0.05, fill: { color: hex(theme.primary) } });
          items.slice(0, 3).forEach((item, i) => {
            const x = 0.6 + i * 3.05;
            slide.addShape(pptx.ShapeType.ellipse, { x: x + 1.15, y: 2.35, w: 0.45, h: 0.45, fill: { color: accentHex } });
            slide.addText(`0${i + 1}`, { x: x + 1.15, y: 2.35, w: 0.45, h: 0.45, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
            slide.addText(item.head, { x, y: 1.65, w: 2.75, h: 0.5, fontSize: Math.round(16 * fScale), bold: true, fontFace: pptxFont, color: hex(theme.primary), align: 'center', shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(12 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x, y: 3.05, w: 2.75, h: 1.85, align: 'center', valign: 'top', shrinkText: true });
          });
        } else if (s.layout === 'metrics') {
          items.slice(0, 3).forEach((item, i) => {
            const x = 0.6 + i * 3.05;
            slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.72, w: 2.75, h: 2.9, rectRadius: 0.1, fill: { color: hex(theme.cardBg) }, line: { color: 'E2E8F0', width: 1 } });
            slide.addShape(pptx.ShapeType.rect, { x, y: 1.72, w: 2.75, h: 0.07, fill: { color: accentHex } });
            slide.addText(item.head, { x, y: 2.05, w: 2.75, h: 1.1, fontSize: Math.round(38 * fScale), bold: true, fontFace: pptxFont, color: accentHex, align: 'center', valign: 'middle', shrinkText: true });
            slide.addText(toPptxRichText(item.body, { fontSize: Math.round(13 * fScale), fontFace: pptxFont, color: hex(theme.text) }, accentHex), { x: x + 0.15, y: 3.25, w: 2.45, h: 1.1, align: 'center', valign: 'top', shrinkText: true });
          });
        }
      }
    }

    pptx.writeFile({ fileName: `${projectName || 'PPT接单成品'}.pptx` });
    showToast('✅ 已导出零弹窗商业级 .pptx（含防溢出排版、原生图表/表格与演讲者备注）！');
  };

  // ================= 公共渲染器：把单页绘制到 1280x720 Canvas =================
  const renderSlideToCanvas = async (ctx, slideData, logoImg, slideIndex, totalSlides) => {
    const isContent = !['cover', 'section', 'ending'].includes(slideData.layout);
    const isDark = !isContent && (theme.isSplit || theme.id === 'minimal');
    const customBgImg = await loadImage(isContent ? contentBg : coverBg);

    if (customBgImg) {
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      const scaleBg = Math.max(1280 / customBgImg.width, 720 / customBgImg.height);
      const drawW = customBgImg.width * scaleBg;
      const drawH = customBgImg.height * scaleBg;
      ctx.drawImage(customBgImg, (1280 - drawW) / 2, (720 - drawH) / 2, drawW, drawH);
    } else if (theme.isSplit && isContent) {
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, 1280, 720);
    } else {
      const grad = ctx.createLinearGradient(0, 0, 1280, 720);
      grad.addColorStop(0, isDark ? theme.coverBgColor || theme.primary : theme.bg);
      grad.addColorStop(1, isDark ? theme.coverBgEnd || '#051526' : theme.bgEnd || theme.bg);
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, 1280, 720);
    }

    if (logoImg) {
      const ratio = Math.min(150 / logoImg.width, 60 / logoImg.height);
      ctx.drawImage(logoImg, 1280 - 76 - logoImg.width * ratio, 32, logoImg.width * ratio, logoImg.height * ratio);
    } else if (headerTag && isContent) {
      ctx.fillStyle = theme.subText;
      ctx.font = `14px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'right';
      ctx.fillText(headerTag, 1204, 48);
      ctx.textAlign = 'left';
    }

    const drawWrappedText = (text, x, y, maxWidth, lineHeight) => {
      const paragraphs = stripMarkdown(text).split('\n');
      let curY = y;
      paragraphs.forEach((para) => {
        let line = '';
        for (const ch of para) {
          if (ctx.measureText(line + ch).width > maxWidth) {
            ctx.fillText(line, x, curY);
            line = ch;
            curY += lineHeight;
          } else {
            line += ch;
          }
        }
        if (line) {
          ctx.fillText(line, x, curY);
          curY += lineHeight;
        }
      });
    };

    const items = slideData.items || [];
    if (showFooter && isContent) {
      ctx.fillStyle = '#CBD5E1';
      ctx.fillRect(76, 662, 1128, 2);
      ctx.fillStyle = theme.subText;
      ctx.font = `14px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'left';
      ctx.fillText(sourceNote || `${projectName || '商业演示方案'}  ·  CONFIDENTIAL`, 76, 690);
      ctx.fillStyle = theme.primary;
      ctx.font = `bold 15px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'right';
      ctx.fillText(`${String(slideIndex + 1).padStart(2, '0')} / ${String(totalSlides).padStart(2, '0')}`, 1204, 690);
    }

    if (!isContent) {
      ctx.textAlign = 'center';
      ctx.fillStyle = isDark ? '#FFFFFF' : theme.text;
      ctx.font = `bold 48px "${fontFamily}", sans-serif`;
      ctx.fillText(slideData.title, 640, 330);
      ctx.fillStyle = isDark ? '#CBD5E1' : theme.subText;
      ctx.font = `24px "${fontFamily}", sans-serif`;
      ctx.fillText(stripMarkdown(slideData.subtitle), 640, 395);
    } else {
      ctx.textAlign = 'left';
      ctx.fillStyle = theme.accent;
      ctx.fillRect(76, 64, 10, 44);
      ctx.fillStyle = theme.text;
      ctx.font = `bold 36px "${fontFamily}", sans-serif`;
      ctx.fillText(slideData.title, 98, 96);
      ctx.fillStyle = theme.subText;
      ctx.font = `18px "${fontFamily}", sans-serif`;
      ctx.fillText(stripMarkdown(slideData.subtitle), 98, 134);

      if (slideData.layout === 'dashboard') {
        const kpis = slideData.kpis && slideData.kpis.length >= 3 ? slideData.kpis : DEFAULT_KPIS;
        kpis.slice(0, 3).forEach((k, i) => {
          const kx = 76 + i * 384;
          const barC = i === 1 ? theme.accent : theme.primary;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(kx, 165, 360, 118);
          ctx.fillStyle = barC;
          ctx.fillRect(kx, 165, 10, 118);
          ctx.fillStyle = theme.subText;
          ctx.font = `16px "${fontFamily}", sans-serif`;
          ctx.fillText(k.label, kx + 24, 196);
          ctx.fillStyle = barC;
          ctx.font = `bold 32px "${fontFamily}", sans-serif`;
          ctx.fillText(k.value, kx + 24, 250);
          if (k.trend) {
            ctx.font = `bold 14px "${fontFamily}", sans-serif`;
            ctx.fillText(k.trend, kx + 200, 246);
          }
        });
        drawChartOnCanvas(ctx, {
          x: 76, y: 305, w: 705, h: 330,
          type: slideData.chartType || 'groupedBar', data: slideData.chartData, theme,
          series1: slideData.chartSeriesName, series2: slideData.chartSeriesName2, unit: slideData.chartUnit, fontFamily,
        });
        items.slice(0, 2).forEach((item, i) => {
          const iy = 305 + i * 172;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(805, iy, 399, 156);
          ctx.fillStyle = i === 0 ? theme.primary : theme.accent;
          ctx.fillRect(805, iy, 399, 6);
          ctx.font = `bold 20px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, 825, iy + 38);
          ctx.fillStyle = theme.text;
          ctx.font = `16px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, 825, iy + 72, 360, 26);
        });
      } else if (slideData.layout === 'chart') {
        drawChartOnCanvas(ctx, {
          x: 76, y: 185, w: 700, h: 440,
          type: slideData.chartType || 'groupedBar', data: slideData.chartData, theme,
          series1: slideData.chartSeriesName, series2: slideData.chartSeriesName2, unit: slideData.chartUnit, fontFamily,
        });
        items.slice(0, 2).forEach((item, i) => {
          const iy = 185 + i * 228;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(804, iy, 400, 210);
          ctx.fillStyle = theme.primary;
          ctx.fillRect(804, iy, 400, 6);
          ctx.font = `bold 22px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, 826, iy + 46);
          ctx.fillStyle = theme.text;
          ctx.font = `18px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, 826, iy + 86, 356, 28);
        });
      } else if (slideData.layout === 'table') {
        const tData = slideData.tableData && slideData.tableData.headers ? slideData.tableData : DEFAULT_TABLE_DATA;
        const colW = 1128 / tData.headers.length;
        const rowH = Math.min(54, Math.floor(300 / (tData.rows.length + 1)));
        tData.headers.forEach((h, ci) => {
          ctx.fillStyle = theme.primary;
          ctx.fillRect(76 + ci * colW, 175, colW - 2, rowH);
          ctx.fillStyle = '#FFFFFF';
          ctx.font = `bold 17px "${fontFamily}", sans-serif`;
          ctx.textAlign = 'center';
          ctx.fillText(h, 76 + ci * colW + colW / 2, 175 + rowH * 0.65);
        });
        tData.rows.forEach((r, ri) => {
          const y = 175 + (ri + 1) * rowH;
          r.forEach((cell, ci) => {
            ctx.fillStyle = ri % 2 === 1 ? '#F1F5F9' : '#FFFFFF';
            ctx.fillRect(76 + ci * colW, y, colW - 2, rowH - 2);
            const isHL = ci === r.length - 1 || /^\+/.test(String(cell));
            ctx.fillStyle = isHL ? theme.accent : theme.text;
            ctx.font = `${isHL || ci === 0 ? 'bold ' : ''}16px "${fontFamily}", sans-serif`;
            ctx.textAlign = ci === 0 ? 'left' : 'center';
            ctx.fillText(String(cell), ci === 0 ? 92 + ci * colW : 76 + ci * colW + colW / 2, y + rowH * 0.65);
          });
        });
        ctx.textAlign = 'left';
        items.slice(0, 2).forEach((item, i) => {
          const ix = 76 + i * 580;
          const barC = i === 0 ? theme.primary : theme.accent;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(ix, 505, 548, 132);
          ctx.fillStyle = barC;
          ctx.fillRect(ix, 505, 10, 132);
          ctx.font = `bold 20px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, ix + 24, 540);
          ctx.fillStyle = theme.text;
          ctx.font = `17px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, ix + 24, 574, 504, 26);
        });
      } else if (slideData.layout === 'fourCol') {
        items.slice(0, 4).forEach((item, i) => {
          const x = 76 + i * 288;
          const headC = i % 2 === 0 ? theme.primary : theme.accent;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(x, 180, 266, 440);
          ctx.fillStyle = headC;
          ctx.fillRect(x, 180, 266, 60);
          ctx.fillStyle = '#FFFFFF';
          ctx.font = `bold 21px "${fontFamily}", sans-serif`;
          ctx.textAlign = 'center';
          ctx.fillText(item.head, x + 133, 218);
          ctx.textAlign = 'left';
          ctx.fillStyle = theme.text;
          ctx.font = `16px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, x + 16, 272, 234, 28);
          if (item.badge) {
            ctx.fillStyle = headC;
            ctx.fillRect(x + 24, 556, 218, 42);
            ctx.fillStyle = '#FFFFFF';
            ctx.font = `bold 17px "${fontFamily}", sans-serif`;
            ctx.textAlign = 'center';
            ctx.fillText(item.badge, x + 133, 583);
            ctx.textAlign = 'left';
          }
        });
      } else if (slideData.layout === 'toc') {
        items.slice(0, 4).forEach((item, i) => {
          const col = i % 2;
          const row = Math.floor(i / 2);
          const x = 76 + col * 580;
          const y = 200 + row * 210;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(x, y, 548, 178);
          ctx.fillStyle = theme.primary;
          ctx.fillRect(x, y + 32, 10, 114);
          ctx.font = `bold 36px "${fontFamily}", sans-serif`;
          ctx.fillText(`0${i + 1}`, x + 32, y + 102);
          ctx.fillStyle = theme.text;
          ctx.font = `bold 22px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, x + 115, y + 72);
          ctx.fillStyle = theme.subText;
          ctx.font = `17px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, x + 115, y + 112, 410, 26);
        });
      } else if (slideData.layout === 'compare') {
        items.slice(0, 2).forEach((item, i) => {
          const x = 76 + i * 580;
          const barC = i === 0 ? theme.subText : theme.primary;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(x, 195, 548, 425);
          ctx.fillStyle = barC;
          ctx.fillRect(x, 195, 548, 8);
          ctx.font = `bold 24px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, x + 28, 250);
          ctx.fillStyle = theme.text;
          ctx.font = `17px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, x + 28, 295, 492, 29);
        });
      } else {
        items.slice(0, 3).forEach((item, i) => {
          const x = 76 + i * 390;
          ctx.fillStyle = theme.cardBg;
          ctx.fillRect(x, 200, 350, 400);
          ctx.fillStyle = theme.primary;
          ctx.fillRect(x, 200, 350, 6);
          ctx.font = `bold 22px "${fontFamily}", sans-serif`;
          ctx.fillText(item.head, x + 24, 255);
          ctx.fillStyle = theme.text;
          ctx.font = `17px "${fontFamily}", sans-serif`;
          drawWrappedText(item.body, x + 24, 305, 300, 28);
        });
      }
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, 1280, 720);
    ctx.clip();
    ctx.translate(640, 360);
    ctx.rotate(-Math.PI / 6);
    ctx.fillStyle = 'rgba(255, 80, 80, 0.24)';
    ctx.font = 'bold 34px sans-serif';
    ctx.textAlign = 'center';
    for (let x = -800; x <= 800; x += 460) {
      for (let y = -500; y <= 500; y += 180) {
        ctx.fillText(watermarkText || '初稿预览 严禁商用', x, y);
      }
    }
    ctx.restore();
  };

  const handleExportWatermarkPNG = async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 1280;
    canvas.height = 720;
    const ctx = canvas.getContext('2d');
    const logoImg = await loadImage(brandLogo);
    await renderSlideToCanvas(ctx, currentSlide, logoImg, currentIdx, slides.length);
    const link = document.createElement('a');
    link.download = `${projectName || '初稿'}-第${currentIdx + 1}页-水印图.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  };

  const handleExportWeChatLongImage = async () => {
    setIsExportingLong(true);
    try {
      const canvasW = 1080;
      const padX = 40;
      const slideW = canvasW - padX * 2;
      const slideH = Math.round((slideW * 9) / 16);
      const headerH = 120;
      const itemLabelH = 38;
      const gapY = 34;
      const footerH = 60;
      const totalH = headerH + slides.length * (itemLabelH + slideH + gapY) + footerH;

      const canvas = document.createElement('canvas');
      canvas.width = canvasW;
      canvas.height = totalH;
      const ctx = canvas.getContext('2d');

      ctx.fillStyle = '#0f172a';
      ctx.fillRect(0, 0, canvasW, totalH);
      ctx.fillStyle = '#1e293b';
      ctx.fillRect(0, 0, canvasW, headerH - 16);
      ctx.fillStyle = '#38bdf8';
      ctx.font = `bold 30px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'left';
      ctx.fillText(`${projectName || 'PPT 演示方案'} · 全套初稿审稿图`, padX, 48);
      ctx.fillStyle = '#94a3b8';
      ctx.font = `16px "${fontFamily}", sans-serif`;
      ctx.fillText(`共 ${slides.length} 页  |  16:9 宽屏标准  |  ${watermarkText || '付尾款后交付无水印可编辑 .pptx'}`, padX, 82);

      const logoImg = await loadImage(brandLogo);
      const scale = slideW / 1280;

      for (let i = 0; i < slides.length; i++) {
        const s = slides[i];
        const blockTop = headerH + i * (itemLabelH + slideH + gapY);
        ctx.fillStyle = '#cbd5e1';
        ctx.font = `bold 18px "${fontFamily}", sans-serif`;
        ctx.textAlign = 'left';
        ctx.fillText(`P${String(i + 1).padStart(2, '0')} / P${String(slides.length).padStart(2, '0')}  ·  ${s.title || '未命名页面'}`, padX, blockTop + 24);
        ctx.save();
        ctx.translate(padX, blockTop + itemLabelH);
        ctx.scale(scale, scale);
        await renderSlideToCanvas(ctx, s, logoImg, i, slides.length);
        ctx.restore();
      }

      const link = document.createElement('a');
      link.download = `${projectName || 'PPT方案'}-全套${slides.length}页微信审稿长图.jpg`;
      link.href = canvas.toDataURL('image/jpeg', 0.92);
      link.click();
      showToast(`📱 已生成全套 ${slides.length} 页微信审稿长图（JPG）！`);
    } finally {
      setIsExportingLong(false);
    }
  };

  const visibleItemCount = getVisibleItemCount(currentSlide.layout);

  return (
    <div style={{ display: 'flex', height: '100vh', background: '#0f172a', color: '#f8fafc', fontFamily: 'sans-serif', overflow: 'hidden' }}>
      {/* 左侧栏 */}
      <div style={{ width: 280, background: '#1e293b', borderRight: '1px solid #334155', display: 'flex', flexDirection: 'column', padding: 14, gap: 10, overflowY: 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ fontWeight: 'bold', fontSize: 16, color: '#38bdf8' }}>PPT 接单工作台 Pro</div>
          <span style={{ fontSize: 10, color: saveStatus.includes('⚠️') ? '#fbbf24' : '#4ade80' }}>{saveStatus}</span>
        </div>

        <button onClick={() => setShowDocModal(true)} style={{ padding: '9px', background: 'linear-gradient(90deg, #2563eb, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, cursor: 'pointer', fontSize: 13, fontWeight: 'bold' }}>
          ✨ 文档一键生成 PPT
        </button>
        <button onClick={() => setShowTemplateModal(true)} style={{ padding: '8px', background: '#334155', color: '#38bdf8', border: '1px solid #38bdf8', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 'bold' }}>
          🎨 全局模板排版与规范中心
        </button>

        <div style={{ padding: 10, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 7 }}>
          <div style={{ fontSize: 12, fontWeight: 'bold', color: '#e2e8f0' }}>客户订单管理（自动防丢）</div>
          <input value={projectName} onChange={(e) => setProjectName(e.target.value)} placeholder="输入客户/项目名称" style={{ width: '100%', padding: '5px 8px', background: '#1e293b', color: '#fff', border: '1px solid #475569', borderRadius: 4, fontSize: 12, boxSizing: 'border-box' }} />
          <div style={{ display: 'flex', gap: 5 }}>
            <button onClick={handleExportProjectJSON} style={{ flex: 1, padding: '5px 4px', fontSize: 11, background: '#0f766e', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>导出订单</button>
            <button onClick={() => jsonInputRef.current.click()} style={{ flex: 1, padding: '5px 4px', fontSize: 11, background: '#334155', color: '#e2e8f0', border: 'none', borderRadius: 4, cursor: 'pointer' }}>导入订单</button>
            <button onClick={handleResetProject} style={{ padding: '5px 7px', fontSize: 11, background: '#475569', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>新建</button>
            <input ref={jsonInputRef} type="file" accept=".json" onChange={handleImportProjectJSON} style={{ display: 'none' }} />
          </div>
        </div>

        <div style={{ padding: 10, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 7 }}>
          <div style={{ fontSize: 12, fontWeight: 'bold', color: '#e2e8f0' }}>客户品牌定制（全局生效）</div>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <button onClick={() => logoInputRef.current.click()} style={{ flex: 1, padding: '6px 8px', fontSize: 11, background: '#1d4ed8', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>{brandLogo ? '更换客户 Logo' : '+ 上传客户 Logo'}</button>
            {brandLogo && <button onClick={() => setBrandLogo('')} style={{ padding: '6px 8px', fontSize: 11, background: '#7f1d1d', color: '#fca5a5', border: 'none', borderRadius: 4, cursor: 'pointer' }}>清除</button>}
            <input ref={logoInputRef} type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: 'none' }} />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: 11, color: '#94a3b8' }}>
            <span>自定义品牌主色：</span>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="color" value={theme.primary} onChange={(e) => setCustomPrimary(e.target.value)} style={{ width: 28, height: 22, border: 'none', background: 'transparent', cursor: 'pointer' }} />
              {customPrimary && <button onClick={() => setCustomPrimary('')} style={{ fontSize: 10, padding: '2px 5px', background: '#334155', color: '#cbd5e1', border: 'none', borderRadius: 3, cursor: 'pointer' }}>重置</button>}
            </div>
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: '#cbd5e1', cursor: 'pointer' }}>
            <input type="checkbox" checked={showFooter} onChange={(e) => setShowFooter(e.target.checked)} /> 开启商业页脚与自动页码
          </label>
        </div>

        <div>
          <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 5 }}>基础配色：</div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5 }}>
            {THEMES.map((t) => (
              <button key={t.id} onClick={() => { setThemeId(t.id); setCustomPrimary(''); }} style={{ padding: '5px 8px', fontSize: 11, borderRadius: 4, cursor: 'pointer', background: themeId === t.id ? t.primary : '#334155', color: '#fff', border: themeId === t.id ? '1px solid #38bdf8' : 'none' }}>
                {t.name}
              </button>
            ))}
          </div>
        </div>

        <label style={{ fontSize: 12, color: '#94a3b8' }}>
          全局字体：
          <select value={fontFamily} onChange={(e) => setFontFamily(e.target.value)} style={{ width: '100%', marginTop: 4, padding: 6, background: '#0f172a', color: '#fff', border: '1px solid #475569', borderRadius: 4, fontSize: 12 }}>
            {SAFE_FONTS.map((f) => <option key={f.value} value={f.value}>{f.name}</option>)}
          </select>
        </label>

        <div style={{ borderTop: '1px solid #334155', paddingTop: 10 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <span style={{ fontSize: 12, color: '#94a3b8' }}>幻灯片列表 ({slides.length}页)</span>
            <button onClick={() => addSlide('dashboard')} style={{ fontSize: 11, padding: '3px 8px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>+ 加一页</button>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
            {slides.map((s, idx) => {
              const audit = analyzeSlideCompliance(s);
              const hasDanger = audit.warnings.some((w) => w.level === 'danger');
              const hasWarn = audit.warnings.length > 0;
              return (
                <div key={s.id} onClick={() => setCurrentIdx(idx)} style={{ padding: '7px 8px', borderRadius: 6, cursor: 'pointer', background: idx === currentIdx ? '#334155' : '#0f172a', border: idx === currentIdx ? '1px solid #38bdf8' : '1px solid #1e293b', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 4 }}>
                  <div style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }}>{idx + 1}. {s.title || '未命名'}</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                    <span style={{ fontSize: 10 }}>{hasDanger ? '🔴' : hasWarn ? '🟡' : '🟢'}</span>
                    <button onClick={(e) => { e.stopPropagation(); moveSlide(idx, -1); }} disabled={idx === 0} style={{ background: 'transparent', border: 'none', color: idx === 0 ? '#475569' : '#94a3b8', cursor: 'pointer', fontSize: 11, padding: '0 2px' }}>↑</button>
                    <button onClick={(e) => { e.stopPropagation(); moveSlide(idx, 1); }} disabled={idx === slides.length - 1} style={{ background: 'transparent', border: 'none', color: idx === slides.length - 1 ? '#475569' : '#94a3b8', cursor: 'pointer', fontSize: 11, padding: '0 2px' }}>↓</button>
                    <button onClick={(e) => { e.stopPropagation(); duplicateSlide(idx); }} style={{ background: 'transparent', border: 'none', color: '#38bdf8', cursor: 'pointer', fontSize: 11, padding: '0 2px' }}>❐</button>
                    <button onClick={(e) => { e.stopPropagation(); deleteSlide(idx); }} style={{ background: 'transparent', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: 12, padding: '0 2px' }}>×</button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 中间栏：实时预览 + 顶部工具条 */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: 18, gap: 10, overflowY: 'auto' }}>
        {toastMsg && <div style={{ background: '#78350f', border: '1px solid #f59e0b', color: '#fef3c7', padding: '9px 14px', borderRadius: 8, fontSize: 12, fontWeight: 'bold' }}>{toastMsg}</div>}

        <div style={{ background: '#1e293b', padding: '10px 14px', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 6 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 12, flexWrap: 'wrap', gap: 8 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={handleUndo} disabled={past.length === 0} style={{ padding: '3px 8px', fontSize: 11, background: past.length ? '#334155' : '#0f172a', color: past.length ? '#fff' : '#475569', border: '1px solid #475569', borderRadius: 4, cursor: past.length ? 'pointer' : 'default' }}>↩️ 撤销</button>
              <button onClick={handleRedo} disabled={future.length === 0} style={{ padding: '3px 8px', fontSize: 11, background: future.length ? '#334155' : '#0f172a', color: future.length ? '#fff' : '#475569', border: '1px solid #475569', borderRadius: 4, cursor: future.length ? 'pointer' : 'default' }}>↪️ 重做</button>
              <span style={{ color: currentAudit.charCount > maxCharLimit ? '#fbbf24' : '#4ade80' }}>字数 {currentAudit.charCount}/{maxCharLimit}</span>
              <div style={{ display: 'inline-flex', background: '#0f172a', borderRadius: 6, padding: 2, border: '1px solid #334155' }}>
                <button onClick={() => setDensityMode('report')} style={{ padding: '2px 8px', fontSize: 11, border: 'none', borderRadius: 4, cursor: 'pointer', background: densityMode === 'report' ? '#0284c7' : 'transparent', color: '#fff' }}>高密汇报(380)</button>
                <button onClick={() => setDensityMode('speech')} style={{ padding: '2px 8px', fontSize: 11, border: 'none', borderRadius: 4, cursor: 'pointer', background: densityMode === 'speech' ? '#0284c7' : 'transparent', color: '#fff' }}>极简演讲(120)</button>
              </div>
              <button onClick={() => setShowAuditModal(true)} style={{ padding: '3px 10px', fontSize: 11, background: totalIssueCount > 0 ? '#92400e' : '#065f46', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer', fontWeight: 'bold' }}>
                🛡️ 全套体检 ({totalIssueCount})
              </button>
            </div>
            <label style={{ cursor: 'pointer', color: '#94a3b8' }}>
              <input type="checkbox" checked={showSafeGuide} onChange={(e) => setShowSafeGuide(e.target.checked)} /> 安全边距线
            </label>
          </div>
          {currentAudit.warnings.map((w, i) => (
            <div key={i} style={{ fontSize: 12, padding: '5px 10px', borderRadius: 4, background: w.level === 'danger' ? '#450a0a' : w.level === 'warn' ? '#422006' : '#172554', color: w.level === 'danger' ? '#fca5a5' : w.level === 'warn' ? '#fde047' : '#93c5fd' }}>
              {w.level === 'danger' ? '🚨 ' : w.level === 'warn' ? '⚠️ ' : '💡 '}{w.msg}
            </div>
          ))}
        </div>

        {/* 16:9 实时预览画布 */}
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div
            style={{
              width: '100%', maxWidth: 900, aspectRatio: '16 / 9',
              background: (isContentSlide ? contentBg : coverBg) ? `url(${isContentSlide ? contentBg : coverBg}) center/cover no-repeat` : canvasBackgroundCss,
              color: activeText, fontFamily: `"${fontFamily}", sans-serif`,
              borderRadius: 10, boxShadow: '0 24px 60px rgba(0,0,0,0.55)', position: 'relative',
              padding: '5% 5.5% 6.5% 5.5%', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', justifyContent: 'center', overflow: 'hidden',
            }}
          >
            {showSafeGuide && <div style={{ position: 'absolute', inset: '5% 5.5%', border: '1px dashed rgba(148,163,184,0.28)', pointerEvents: 'none' }} />}
            {brandLogo ? (
              <img src={brandLogo} alt="Logo" style={{ position: 'absolute', top: '4%', right: '5.5%', maxHeight: '8%', maxWidth: '14%', objectFit: 'contain' }} />
            ) : (
              headerTag && isContentSlide && (
                <div style={{ position: 'absolute', top: '4.2%', right: '5.5%', fontSize: 10, color: theme.subText }}>{headerTag}</div>
              )
            )}

            {currentSlide.layout === 'cover' && (
              <div style={{ borderLeft: `6px solid ${isDarkCover ? theme.accent : theme.primary}`, paddingLeft: 26 }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 4, color: isDarkCover ? theme.accent : theme.primary, marginBottom: 10 }}>ANNUAL REPORT / PRESENTATION</div>
                <div style={{ fontSize: Math.round(35 * currentFontScale), fontWeight: 900, lineHeight: 1.25 }}>{currentSlide.title}</div>
                <div style={{ fontSize: Math.round(16 * currentFontScale), color: activeSubText, marginTop: 14 }}>{currentSlide.subtitle}</div>
              </div>
            )}

            {currentSlide.layout === 'section' && (
              <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                <div style={{ fontSize: 22, fontWeight: 900, color: isDarkCover ? theme.accent : theme.primary, letterSpacing: 3 }}>{currentSlide.subtitle || 'PART 01'}</div>
                <div style={{ width: 64, height: 4, background: theme.accent, margin: '14px 0', borderRadius: 2 }} />
                <div style={{ fontSize: Math.round(36 * currentFontScale), fontWeight: 900 }}>{currentSlide.title}</div>
              </div>
            )}

            {currentSlide.layout === 'ending' && (
              <div style={{ textAlign: 'center' }}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 4, color: activeSubText, marginBottom: 10 }}>THANK YOU FOR WATCHING</div>
                <div style={{ fontSize: Math.round(40 * currentFontScale), fontWeight: 900, color: isDarkCover ? '#FFFFFF' : theme.primary }}>{currentSlide.title}</div>
                <div style={{ fontSize: 16, color: activeSubText, marginTop: 14 }}>{currentSlide.subtitle}</div>
              </div>
            )}

            {isContentSlide && (
              <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, paddingRight: brandLogo || headerTag ? '20%' : 0 }}>
                  <div style={{ width: 6, height: 24, background: theme.accent, borderRadius: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: Math.round(20 * currentFontScale), fontWeight: 900 }}>{currentSlide.title}</div>
                </div>
                <div style={{ fontSize: 11.5, color: theme.subText, marginTop: 3, marginBottom: 12, paddingLeft: 16 }}>{renderRichText(currentSlide.subtitle, theme.accent)}</div>

                {currentSlide.layout === 'dashboard' && (
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 12, minHeight: 0 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12 }}>
                      {currentKpis.slice(0, 3).map((k, idx) => {
                        const barColor = idx === 1 ? theme.accent : theme.primary;
                        return (
                          <div key={idx} style={{ background: theme.cardBg, borderLeft: `4px solid ${barColor}`, borderRadius: 8, padding: '8px 12px', border: '1px solid #e2e8f0', boxShadow: cardShadow, display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
                            <div style={{ fontSize: 10.5, color: theme.subText }}>{k.label}</div>
                            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', marginTop: 4 }}>
                              <span style={{ fontSize: 21, fontWeight: 900, color: barColor }}>{k.value}</span>
                              {k.trend && <span style={{ fontSize: 9.5, fontWeight: 700, padding: '2px 7px', borderRadius: 12, background: '#f1f5f9', color: barColor }}>{k.trend}</span>}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <div style={{ flex: 1, display: 'grid', gridTemplateColumns: isReversed ? '1fr 1.65fr' : '1.65fr 1fr', gap: 14, minHeight: 0 }}>
                      <div style={{ order: isReversed ? 2 : 1, background: theme.cardBg, borderRadius: 8, padding: 10, border: '1px solid #e2e8f0', boxShadow: cardShadow, display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden' }}>
                        <ChartPreview type={currentChartType} data={currentChartData} theme={theme} series1={currentSlide.chartSeriesName} series2={currentSlide.chartSeriesName2} unit={currentSlide.chartUnit} />
                      </div>
                      <div style={{ order: isReversed ? 1 : 2, display: 'flex', flexDirection: 'column', gap: 10 }}>
                        {(currentSlide.items || []).slice(0, 2).map((item, idx) => (
                          <div key={idx} style={{ flex: 1, background: theme.cardBg, borderTop: `3px solid ${idx === 0 ? theme.primary : theme.accent}`, borderRadius: 8, padding: '9px 11px', border: '1px solid #e2e8f0', boxShadow: cardShadow, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                            <div style={{ fontSize: Math.round(12.5 * currentFontScale), fontWeight: 800, color: idx === 0 ? theme.primary : theme.accent, marginBottom: 3 }}>{item.head}</div>
                            <div style={{ fontSize: Math.round(11 * currentFontScale), lineHeight: 1.45, color: theme.text }}>{renderRichText(item.body, theme.accent)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                )}

                {currentSlide.layout === 'table' && (
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 10, minHeight: 0 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: Math.round(11.5 * currentFontScale), background: '#FFFFFF', borderRadius: 6, overflow: 'hidden', boxShadow: cardShadow }}>
                      <thead>
                        <tr>
                          {currentTableData.headers.map((h, ci) => (
                            <th key={ci} style={{ background: theme.primary, color: '#FFFFFF', padding: '7px 9px', textAlign: ci === 0 ? 'left' : 'center', fontWeight: 700, border: '1px solid #cbd5e1' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {currentTableData.rows.map((row, ri) => (
                          <tr key={ri} style={{ background: ri % 2 === 1 ? '#f1f5f9' : '#ffffff' }}>
                            {row.map((cell, ci) => {
                              const isHL = ci === row.length - 1 || /^\+/.test(String(cell).trim());
                              return <td key={ci} style={{ padding: '6px 9px', textAlign: ci === 0 ? 'left' : 'center', color: isHL ? theme.accent : '#0f172a', fontWeight: isHL || ci === 0 ? 700 : 400, border: '1px solid #e2e8f0' }}>{cell}</td>;
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginTop: 'auto' }}>
                      {(currentSlide.items || []).slice(0, 2).map((item, idx) => (
                        <div key={idx} style={{ background: theme.cardBg, borderLeft: `4px solid ${idx === 0 ? theme.primary : theme.accent}`, borderRadius: 6, padding: '8px 11px', border: '1px solid #e2e8f0' }}>
                          <div style={{ fontSize: Math.round(12 * currentFontScale), fontWeight: 800, color: idx === 0 ? theme.primary : theme.accent }}>{item.head}</div>
                          <div style={{ fontSize: Math.round(11 * currentFontScale), color: theme.text, marginTop: 3, lineHeight: 1.4 }}>{renderRichText(item.body, theme.accent)}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {currentSlide.layout === 'fourCol' && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, minHeight: 0 }}>
                    {(currentSlide.items || []).slice(0, 4).map((item, idx) => {
                      const headColor = idx % 2 === 0 ? theme.primary : theme.accent;
                      return (
                        <div key={idx} style={{ background: theme.cardBg, borderRadius: 8, border: '1px solid #e2e8f0', boxShadow: cardShadow, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                          <div style={{ background: headColor, color: '#FFFFFF', padding: '8px 6px', textAlign: 'center', fontSize: Math.round(13 * currentFontScale), fontWeight: 800 }}>{item.head}</div>
                          <div style={{ flex: 1, padding: '9px 9px', fontSize: Math.round(11 * currentFontScale), lineHeight: 1.55, color: theme.text, whiteSpace: 'pre-line', overflow: 'hidden' }}>{renderRichText(item.body, theme.accent)}</div>
                          {item.badge && (
                            <div style={{ margin: '0 10px 9px', padding: '4px 6px', background: headColor, color: '#FFFFFF', borderRadius: 20, textAlign: 'center', fontSize: 10.5, fontWeight: 700 }}>{item.badge}</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}

                {currentSlide.layout === 'chart' && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: isReversed ? '1fr 1.65fr' : '1.65fr 1fr', gap: 16, minHeight: 0 }}>
                    <div style={{ order: isReversed ? 2 : 1, background: theme.cardBg, borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', boxShadow: cardShadow, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      <ChartPreview type={currentChartType} data={currentChartData} theme={theme} series1={currentSlide.chartSeriesName} series2={currentSlide.chartSeriesName2} unit={currentSlide.chartUnit} />
                    </div>
                    <div style={{ order: isReversed ? 1 : 2, display: 'flex', flexDirection: 'column', gap: 12, justifyContent: 'center' }}>
                      {(currentSlide.items || []).slice(0, 2).map((item, idx) => (
                        <div key={idx} style={{ flex: 1, background: theme.cardBg, borderTop: `4px solid ${theme.primary}`, borderRadius: 8, padding: 12, boxShadow: cardShadow, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                          <div style={{ fontSize: Math.round(13.5 * currentFontScale), fontWeight: 700, color: theme.primary, marginBottom: 4 }}>{item.head}</div>
                          <div style={{ fontSize: Math.round(11.5 * currentFontScale), lineHeight: 1.5, color: theme.text }}>{renderRichText(item.body, theme.accent)}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {currentSlide.layout === 'toc' && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14, alignContent: 'center' }}>
                    {(currentSlide.items || []).slice(0, 4).map((item, idx) => (
                      <div key={idx} style={{ background: theme.cardBg, borderLeft: `4px solid ${theme.primary}`, borderRadius: 8, padding: 14, boxShadow: cardShadow, display: 'flex', gap: 12, alignItems: 'center' }}>
                        <div style={{ fontSize: 26, fontWeight: 900, color: theme.primary }}>0{idx + 1}</div>
                        <div>
                          <div style={{ fontSize: Math.round(15 * currentFontScale), fontWeight: 700 }}>{item.head}</div>
                          <div style={{ fontSize: Math.round(12 * currentFontScale), color: theme.subText, marginTop: 3 }}>{renderRichText(item.body, theme.accent)}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {currentSlide.layout === 'imageText' && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, alignItems: 'center', minHeight: 0 }}>
                    <div style={{ order: isReversed ? 2 : 1, height: '100%', background: theme.cardBg, borderRadius: 8, boxShadow: cardShadow, display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden' }}>
                      {currentSlide.image ? <img src={currentSlide.image} alt="配图" style={{ width: '100%', height: '100%', objectFit: 'contain' }} /> : <span style={{ fontSize: 13, color: theme.subText }}>请在右侧上传配图</span>}
                    </div>
                    <div style={{ order: isReversed ? 1 : 2, display: 'flex', flexDirection: 'column', gap: 14 }}>
                      {(currentSlide.items || []).slice(0, 3).map((item, idx) => (
                        <div key={idx} style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                          <div style={{ width: 32, height: 32, borderRadius: 8, background: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                            <IconSvg id={item.icon || ICONS[idx].id} color={theme.primary} size={18} />
                          </div>
                          <div>
                            <div style={{ fontSize: Math.round(14 * currentFontScale), fontWeight: 700, color: theme.text }}>{item.head}</div>
                            <div style={{ fontSize: Math.round(11.5 * currentFontScale), marginTop: 3, lineHeight: 1.5, color: theme.subText }}>{renderRichText(item.body, theme.accent)}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {currentSlide.layout === 'compare' && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 18, alignItems: 'stretch' }}>
                    {(currentSlide.items || []).slice(0, 2).map((item, idx) => (
                      <div key={idx} style={{ background: theme.cardBg, borderTop: `4px solid ${idx === 0 ? theme.subText : theme.primary}`, borderRadius: 8, padding: 16, boxShadow: cardShadow, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                        <div style={{ fontSize: Math.round(16 * currentFontScale), fontWeight: 900, color: idx === 0 ? theme.subText : theme.primary, paddingBottom: 6 }}>{item.head}</div>
                        <div style={{ fontSize: Math.round(11.5 * currentFontScale), lineHeight: 1.65, marginTop: 4, whiteSpace: 'pre-line' }}>{renderRichText(item.body, theme.accent)}</div>
                      </div>
                    ))}
                  </div>
                )}

                {['cards', 'timeline', 'metrics'].includes(currentSlide.layout) && (
                  <div style={{ flex: 1, display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, alignItems: 'center' }}>
                    {(currentSlide.items || []).slice(0, 3).map((item, idx) => (
                      <div key={idx} style={{ background: currentSlide.layout === 'timeline' ? 'transparent' : theme.cardBg, borderTop: currentSlide.layout === 'cards' ? `4px solid ${theme.primary}` : currentSlide.layout === 'metrics' ? `4px solid ${theme.accent}` : 'none', boxShadow: currentSlide.layout === 'timeline' ? 'none' : cardShadow, borderRadius: 8, padding: 14, height: '88%', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', justifyContent: currentSlide.layout === 'metrics' ? 'center' : 'flex-start', textAlign: currentSlide.layout === 'cards' ? 'left' : 'center', overflow: 'hidden' }}>
                        {currentSlide.layout === 'cards' && (
                          <div style={{ width: 34, height: 34, borderRadius: 8, background: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 10 }}>
                            <IconSvg id={item.icon || ICONS[idx].id} color={theme.primary} size={18} />
                          </div>
                        )}
                        {currentSlide.layout === 'timeline' && (
                          <div style={{ width: 28, height: 28, borderRadius: '50%', background: theme.accent, color: '#fff', fontWeight: 'bold', lineHeight: '28px', margin: '0 auto 10px', fontSize: 13 }}>0{idx + 1}</div>
                        )}
                        <div style={{ fontSize: currentSlide.layout === 'metrics' ? Math.round(32 * currentFontScale) : Math.round(15 * currentFontScale), fontWeight: 900, color: currentSlide.layout === 'metrics' ? theme.accent : theme.text, marginBottom: 6 }}>{item.head}</div>
                        <div style={{ fontSize: Math.round(11 * currentFontScale), lineHeight: 1.55, color: theme.subText, whiteSpace: 'pre-line' }}>{renderRichText(item.body, theme.accent)}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {showFooter && isContentSlide && (
              <div style={{ position: 'absolute', left: '5.5%', right: '5.5%', bottom: '2.5%', borderTop: '1px solid #cbd5e1', paddingTop: 5, display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 10 }}>
                <span style={{ color: theme.subText }}>{sourceNote || `${projectName || '商业演示方案'} · CONFIDENTIAL`}</span>
                <span style={{ color: theme.primary, fontWeight: 'bold' }}>{String(currentIdx + 1).padStart(2, '0')} / {String(slides.length).padStart(2, '0')}</span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* 右侧栏：填空编辑 + 单页微调 + Excel粘贴 + 演讲稿备注 + 导出 */}
      <div style={{ width: 350, background: '#1e293b', borderLeft: '1px solid #334155', padding: 14, display: 'flex', flexDirection: 'column', gap: 10, overflowY: 'auto' }}>
        <div style={{ fontWeight: 'bold', fontSize: 14 }}>当前页填空编辑</div>

        <label style={{ fontSize: 12, color: '#94a3b8' }}>
          切换本页版式：
          <select value={currentSlide.layout} onChange={(e) => updateSlide({ layout: e.target.value })} style={{ width: '100%', marginTop: 4, padding: 6, background: '#0f172a', color: '#fff', border: '1px solid #475569', borderRadius: 4 }}>
            {LAYOUTS.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
        </label>

        {/* 单页排版微调：字号缩放 + 左右镜像 */}
        <div style={{ display: 'flex', gap: 6, alignItems: 'center', padding: '6px 8px', background: '#0f172a', borderRadius: 6, border: '1px solid #334155', fontSize: 11 }}>
          <span style={{ color: '#94a3b8' }}>本页字号：</span>
          {[0.85, 1, 1.1].map((sc) => (
            <button key={sc} onClick={() => updateSlide({ fontScale: sc })} style={{ padding: '2px 7px', fontSize: 11, borderRadius: 4, border: 'none', cursor: 'pointer', background: currentFontScale === sc ? '#0284c7' : '#1e293b', color: '#fff' }}>
              {Math.round(sc * 100)}%
            </button>
          ))}
          {supportsMirror(currentSlide.layout) && (
            <button onClick={() => updateSlide({ reverseLayout: !isReversed })} style={{ marginLeft: 'auto', padding: '2px 8px', fontSize: 11, borderRadius: 4, border: '1px solid #38bdf8', cursor: 'pointer', background: isReversed ? '#0284c7' : 'transparent', color: '#fff' }}>
              ⇄ 左右互换
            </button>
          )}
        </div>

        <label style={{ fontSize: 12, color: '#94a3b8' }}>
          主标题：
          <input value={currentSlide.title} onChange={(e) => updateSlide({ title: e.target.value })} style={{ width: '100%', marginTop: 4, padding: 6, background: '#0f172a', color: '#fff', border: '1px solid #475569', borderRadius: 4, boxSizing: 'border-box' }} />
        </label>

        <label style={{ fontSize: 12, color: '#94a3b8' }}>
          副标题 / 说明（支持 **关键词** 标红）：
          <input value={currentSlide.subtitle} onChange={(e) => updateSlide({ subtitle: e.target.value })} style={{ width: '100%', marginTop: 4, padding: 6, background: '#0f172a', color: '#fff', border: '1px solid #475569', borderRadius: 4, boxSizing: 'border-box' }} />
        </label>

        {currentSlide.layout === 'dashboard' && (
          <div style={{ padding: 10, background: '#0f172a', borderRadius: 6, border: '1px solid #f59e0b', display: 'flex', flexDirection: 'column', gap: 6 }}>
            <div style={{ fontSize: 12, fontWeight: 'bold', color: '#fbbf24' }}>🏆 顶部 3 项核心 KPI 指标</div>
            {currentKpis.slice(0, 3).map((k, idx) => (
              <div key={idx} style={{ display: 'flex', gap: 4 }}>
                <input value={k.label} onChange={(e) => updateKpi(idx, 'label', e.target.value)} placeholder="指标名" style={{ flex: 1.2, minWidth: 0, padding: 4, background: '#1e293b', color: '#cbd5e1', border: '1px solid #334155', borderRadius: 3, fontSize: 11 }} />
                <input value={k.value} onChange={(e) => updateKpi(idx, 'value', e.target.value)} placeholder="数值" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#38bdf8', border: '1px solid #334155', borderRadius: 3, fontSize: 11, fontWeight: 'bold' }} />
                <input value={k.trend} onChange={(e) => updateKpi(idx, 'trend', e.target.value)} placeholder="同比" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#f87171', border: '1px solid #334155', borderRadius: 3, fontSize: 11 }} />
              </div>
            ))}
          </div>
        )}

        {showChartEditor && (
          <div style={{ padding: 10, background: '#0f172a', borderRadius: 6, border: '1px solid #38bdf8', display: 'flex', flexDirection: 'column', gap: 7 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: 12, fontWeight: 'bold', color: '#38bdf8' }}>📊 多系列图表配置</span>
              <button onClick={() => setShowExcelPaste((v) => !v)} style={{ padding: '2px 7px', fontSize: 10, background: '#0f766e', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>📋 粘贴 Excel</button>
            </div>
            {showExcelPaste && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5, padding: 6, background: '#1e293b', borderRadius: 4 }}>
                <textarea rows={3} value={excelRawText} onChange={(e) => setExcelRawText(e.target.value)} placeholder="直接从 Excel 复制 2~3 列粘贴到这里..." style={{ width: '100%', fontSize: 11, background: '#0f172a', color: '#4ade80', border: '1px solid #334155', borderRadius: 4, boxSizing: 'border-box' }} />
                <button onClick={() => handleApplyExcelPaste('chart')} style={{ padding: 4, fontSize: 11, background: '#059669', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>确认导入图表数据</button>
              </div>
            )}
            <select value={currentChartType} onChange={(e) => updateSlide({ chartType: e.target.value })} style={{ width: '100%', padding: 5, background: '#1e293b', color: '#fff', border: '1px solid #475569', borderRadius: 4, fontSize: 12 }}>
              {CHART_TYPES.map((ct) => <option key={ct.id} value={ct.id}>{ct.name}</option>)}
            </select>
            <div style={{ display: 'flex', gap: 5 }}>
              <input value={currentSlide.chartUnit || ''} onChange={(e) => updateSlide({ chartUnit: e.target.value })} placeholder="单位(如 亿元)" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#94a3b8', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
              <input value={currentSlide.chartSeriesName || ''} onChange={(e) => updateSlide({ chartSeriesName: e.target.value })} placeholder="系列1" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#38bdf8', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
              {isMultiSeries(currentChartType) && (
                <input value={currentSlide.chartSeriesName2 ?? ''} onChange={(e) => updateSlide({ chartSeriesName2: e.target.value })} placeholder="系列2(留空隐藏)" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#f87171', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
              )}
            </div>
            {currentChartData.map((row, idx) => (
              <div key={idx} style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
                <input value={row.label} onChange={(e) => updateChartRow(idx, 'label', e.target.value)} placeholder="类目" style={{ flex: 1.3, minWidth: 0, padding: 4, background: '#1e293b', color: '#fff', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
                <input type="number" value={row.value} onChange={(e) => updateChartRow(idx, 'value', e.target.value)} placeholder="值1" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#38bdf8', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
                {isMultiSeries(currentChartType) && (
                  <input type="number" value={row.value2 ?? 0} onChange={(e) => updateChartRow(idx, 'value2', e.target.value)} placeholder="值2" style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#f87171', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }} />
                )}
                <button onClick={() => removeChartRow(idx)} style={{ background: 'transparent', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: 14, padding: '0 2px' }}>×</button>
              </div>
            ))}
            <button onClick={addChartRow} style={{ padding: '4px', fontSize: 11, background: '#1e293b', color: '#38bdf8', border: '1px dashed #38bdf8', borderRadius: 4, cursor: 'pointer' }}>+ 添加一行数据</button>
          </div>
        )}

        {currentSlide.layout === 'table' && (
          <div style={{ padding: 10, background: '#0f172a', borderRadius: 6, border: '1px solid #38bdf8', display: 'flex', flexDirection: 'column', gap: 8 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: 12, fontWeight: 'bold', color: '#38bdf8' }}>📋 表格数据</span>
              <div style={{ display: 'flex', gap: 4 }}>
                <button onClick={() => setShowExcelPaste((v) => !v)} style={{ padding: '2px 6px', fontSize: 10, background: '#059669', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>粘贴Excel</button>
                <button onClick={addTableRow} style={{ padding: '2px 6px', fontSize: 11, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>+行</button>
                <button onClick={addTableCol} style={{ padding: '2px 6px', fontSize: 11, background: '#0f766e', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>+列</button>
                <button onClick={removeTableCol} style={{ padding: '2px 6px', fontSize: 11, background: '#475569', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>-列</button>
              </div>
            </div>
            {showExcelPaste && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5, padding: 6, background: '#1e293b', borderRadius: 4 }}>
                <textarea rows={3} value={excelRawText} onChange={(e) => setExcelRawText(e.target.value)} placeholder="直接从 Excel 复制表格粘贴到这里（首行自动作表头）..." style={{ width: '100%', fontSize: 11, background: '#0f172a', color: '#4ade80', border: '1px solid #334155', borderRadius: 4, boxSizing: 'border-box' }} />
                <button onClick={() => handleApplyExcelPaste('table')} style={{ padding: 4, fontSize: 11, background: '#059669', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>确认导入表格</button>
              </div>
            )}
            <div style={{ display: 'flex', gap: 3 }}>
              {currentTableData.headers.map((h, ci) => (
                <input key={ci} value={h} onChange={(e) => updateTableHeader(ci, e.target.value)} style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e3a5f', color: '#fff', border: '1px solid #334155', borderRadius: 3, fontSize: 11, fontWeight: 'bold' }} />
              ))}
            </div>
            {currentTableData.rows.map((row, ri) => (
              <div key={ri} style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
                {row.map((cell, ci) => (
                  <input key={ci} value={cell} onChange={(e) => updateTableCell(ri, ci, e.target.value)} style={{ flex: 1, minWidth: 0, padding: 4, background: '#1e293b', color: '#fff', border: '1px solid #334155', borderRadius: 3, fontSize: 11 }} />
                ))}
                <button onClick={() => removeTableRow(ri)} style={{ background: 'transparent', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: 13, padding: '0 2px' }}>×</button>
              </div>
            ))}
          </div>
        )}

        {currentSlide.layout === 'imageText' && (
          <div style={{ padding: 10, background: '#0f172a', borderRadius: 6, border: '1px dashed #38bdf8', display: 'flex', flexDirection: 'column', gap: 6 }}>
            <div style={{ fontSize: 12, fontWeight: 'bold', color: '#38bdf8' }}>配图上传（合规防踩坑）</div>
            <div style={{ display: 'flex', gap: 6 }}>
              <button onClick={() => slideImgInputRef.current.click()} style={{ flex: 1, padding: '6px', fontSize: 12, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>{currentSlide.image ? '更换配图' : '+ 上传本地图片'}</button>
              {currentSlide.image && <button onClick={() => updateSlide({ image: '' })} style={{ padding: '6px 10px', fontSize: 12, background: '#7f1d1d', color: '#fca5a5', border: 'none', borderRadius: 4, cursor: 'pointer' }}>清除</button>}
              <input ref={slideImgInputRef} type="file" accept="image/*" onChange={handleSlideImageUpload} style={{ display: 'none' }} />
            </div>
          </div>
        )}

        {visibleItemCount > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 7 }}>
            <div style={{ fontSize: 11, color: '#94a3b8' }}>💡 用 <code style={{ color: '#f87171' }}>**重点**</code> 包裹可自动加粗标红</div>
            {Array.from({ length: visibleItemCount }).map((_, idx) => {
              const item = (currentSlide.items && currentSlide.items[idx]) || { head: '', body: '', icon: ICONS[idx % ICONS.length].id, badge: '' };
              return (
                <div key={idx} style={{ padding: 8, background: '#0f172a', borderRadius: 6, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 5 }}>
                  <div style={{ display: 'flex', gap: 5 }}>
                    {supportsIcons(currentSlide.layout) && (
                      <select value={item.icon || ICONS[idx % ICONS.length].id} onChange={(e) => updateItem(idx, 'icon', e.target.value)} style={{ width: 72, padding: 4, background: '#1e293b', color: '#38bdf8', border: '1px solid #334155', borderRadius: 4, fontSize: 11 }}>
                        {ICONS.map((ic) => <option key={ic.id} value={ic.id}>{ic.name}</option>)}
                      </select>
                    )}
                    <input value={item.head} onChange={(e) => updateItem(idx, 'head', e.target.value)} placeholder="小标题 / 结论标题" style={{ flex: 1, padding: 5, background: '#1e293b', color: '#38bdf8', border: '1px solid #334155', borderRadius: 4, fontSize: 12, boxSizing: 'border-box' }} />
                  </div>
                  <textarea rows={currentSlide.layout === 'fourCol' || currentSlide.layout === 'compare' ? 3 : 2} value={item.body} onChange={(e) => updateItem(idx, 'body', e.target.value)} placeholder="支持回车分点与 **标红**" style={{ width: '100%', padding: 5, background: '#1e293b', color: '#fff', border: '1px solid #334155', borderRadius: 4, fontSize: 12, boxSizing: 'border-box' }} />
                  {currentSlide.layout === 'fourCol' && (
                    <input value={item.badge || ''} onChange={(e) => updateItem(idx, 'badge', e.target.value)} placeholder="底部数据胶囊（如 GMV 3.25亿）" style={{ width: '100%', padding: 4, background: '#1e293b', color: '#fbbf24', border: '1px solid #334155', borderRadius: 4, fontSize: 11, boxSizing: 'border-box' }} />
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* 演讲者备注 / 答辩逐字稿 */}
        <label style={{ fontSize: 11, color: '#94a3b8' }}>
          🎤 本页演讲者备注 / 答辩逐字稿（原生写入 PPT 备注栏）：
          <textarea rows={2} value={currentSlide.notes || ''} onChange={(e) => updateSlide({ notes: e.target.value })} placeholder="输入本页口播讲稿（选填）..." style={{ width: '100%', marginTop: 4, padding: 5, background: '#0f172a', color: '#cbd5e1', border: '1px solid #334155', borderRadius: 4, fontSize: 11, boxSizing: 'border-box' }} />
        </label>

        <div style={{ marginTop: 'auto', borderTop: '1px solid #334155', paddingTop: 10, display: 'flex', flexDirection: 'column', gap: 7 }}>
          <input value={watermarkText} onChange={(e) => setWatermarkText(e.target.value)} placeholder="防跑单水印文字" style={{ width: '100%', padding: 5, background: '#0f172a', color: '#fca5a5', border: '1px solid #475569', borderRadius: 4, fontSize: 11, boxSizing: 'border-box' }} />
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={handleExportWatermarkPNG} style={{ flex: 1, padding: '8px 6px', background: '#475569', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 12 }}>单页水印图</button>
            <button onClick={handleExportWeChatLongImage} disabled={isExportingLong} style={{ flex: 1.4, padding: '8px 6px', background: '#059669', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 12, fontWeight: 'bold' }}>
              {isExportingLong ? '拼图中...' : '📱 导出全套微信长图'}
            </button>
          </div>
          <button onClick={handleExportPPTX} style={{ padding: 11, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 14, fontWeight: 'bold' }}>
            导出可编辑成品 (.pptx)
          </button>
        </div>
      </div>

      {/* 全套一键交付体检弹窗 */}
      {showAuditModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(2, 6, 23, 0.8)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 20 }}>
          <div style={{ width: '100%', maxWidth: 640, maxHeight: '85vh', overflowY: 'auto', background: '#1e293b', border: '1px solid #38bdf8', borderRadius: 12, padding: 20, display: 'flex', flexDirection: 'column', gap: 12, color: '#f8fafc' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ fontSize: 16, fontWeight: 'bold', color: '#38bdf8' }}>🛡️ 全套交付前合规与排版体检（共 {slides.length} 页）</div>
              <button onClick={() => setShowAuditModal(false)} style={{ background: '#334155', color: '#fff', border: 'none', borderRadius: 6, padding: '5px 12px', cursor: 'pointer' }}>关闭</button>
            </div>
            {totalIssueCount === 0 ? (
              <div style={{ padding: 24, textAlign: 'center', background: '#052e16', color: '#4ade80', borderRadius: 8, fontWeight: 'bold' }}>
                ✅ 全部 {slides.length} 页均已通过字数、敏感词与占位符检查，可直接交付！
              </div>
            ) : (
              fullDeckAudits.filter((a) => a.warnings.length > 0).map((a) => (
                <div key={a.idx} style={{ padding: 10, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 6 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <strong style={{ fontSize: 13 }}>第 {a.idx + 1} 页：{a.title || '未命名'}</strong>
                    <button onClick={() => { setCurrentIdx(a.idx); setShowAuditModal(false); }} style={{ padding: '3px 10px', fontSize: 11, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>去修改 →</button>
                  </div>
                  {a.warnings.map((w, wi) => (
                    <div key={wi} style={{ fontSize: 11.5, color: w.level === 'danger' ? '#fca5a5' : w.level === 'warn' ? '#fde047' : '#93c5fd' }}>
                      {w.level === 'danger' ? '🚨 ' : w.level === 'warn' ? '⚠️ ' : '💡 '}{w.msg}
                    </div>
                  ))}
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {showTemplateModal && (
        <TemplateModal
          config={{ themeId, customPrimary, fontFamily, brandLogo, showFooter, coverBg, contentBg, headerTag, sourceNote }}
          onClose={() => setShowTemplateModal(false)}
          onApply={(patch, msg) => {
            if (patch.themeId !== undefined) setThemeId(patch.themeId);
            if (patch.customPrimary !== undefined) setCustomPrimary(patch.customPrimary);
            if (patch.fontFamily !== undefined) setFontFamily(patch.fontFamily);
            if (patch.brandLogo !== undefined) setBrandLogo(patch.brandLogo);
            if (patch.showFooter !== undefined) setShowFooter(patch.showFooter);
            if (patch.coverBg !== undefined) setCoverBg(patch.coverBg);
            if (patch.contentBg !== undefined) setContentBg(patch.contentBg);
            if (patch.headerTag !== undefined) setHeaderTag(patch.headerTag);
            if (patch.sourceNote !== undefined) setSourceNote(patch.sourceNote);
            if (msg) showToast(msg);
          }}
        />
      )}

      {showDocModal && (
        <DocToPptModal
          onClose={() => setShowDocModal(false)}
          onGenerate={(newSlides, newName) => {
            commitSlides(newSlides);
            setProjectName(newName);
            setCurrentIdx(0);
            setShowDocModal(false);
            showToast(`🚀 已根据文档自动生成 ${newSlides.length} 页高密汇报 PPT，请逐页精修！`);
          }}
        />
      )}
    </div>
  );
}