import { useState, useRef } from 'react';
import mammoth from 'mammoth';

const VALID_LAYOUTS = ['cover', 'toc', 'section', 'dashboard', 'table', 'fourCol', 'chart', 'compare', 'imageText', 'cards', 'timeline', 'metrics', 'ending'];
const VALID_ICONS = ['trend', 'rocket', 'shield', 'target', 'zap', 'award', 'users', 'bulb', 'layers', 'clock', 'chart', 'globe'];
const VALID_CHARTS = ['groupedBar', 'line', 'doughnut', 'bar', 'pie'];

export default function DocToPptModal({ onClose, onGenerate }) {
  const [rawText, setRawText] = useState('');
  const [fileName, setFileName] = useState('');
  const [slideCount, setSlideCount] = useState(10);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const fileInputRef = useRef(null);

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name.replace(/\.(docx|txt|md)$/i, ''));
    try {
      if (/\.docx$/i.test(file.name)) {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        setRawText(result.value.trim());
      } else {
        const text = await file.text();
        setRawText(text.trim());
      }
    } catch {
      alert('文档读取失败，请确认是标准的 .docx / .txt / .md 文件（暂不支持老版 .doc）。');
    }
    e.target.value = '';
  };

  const handleCopyPrompt = async () => {
    if (!rawText.trim()) return alert('请先上传 Word/TXT 文档，或在下方粘贴文案！');

    const prompt = `你是一位顶级大厂/咨询风商业 PPT 策划师（旁门左道汇报风格）。请将以下文档拆解为一套约 ${slideCount} 页的高密工作汇报 PPT，严格按下方 JSON 格式输出（只输出纯 JSON 代码块，不要多余解释）。

【排版与高密汇报规范】
1. 必须包含封面(cover)、目录(toc)、结尾(ending)，中间页优先使用高密汇报版式：
   - dashboard（经营仪表盘：顶部3个KPI + 下方多系列图表 + 右侧2条归因）
   - table（商务数据表格：表头+多行数据 + 下方2条结论）
   - fourCol（四栏/四季度：顶部标题块 + 多行要点 + 底部数据胶囊 badge）
   - chart（数据图表页，chartType 可选 groupedBar双柱对比 / line折线 / doughnut环形 / bar单柱 / pie饼图）
   - compare（左右对比）、cards（三栏图标卡片）、timeline（时间轴）、metrics（核心大数字）、imageText（左图右文）、section（章节过渡）。
2. 正文与卡片支持用 **关键词** 包裹核心数字或结论（系统会自动加粗标红），单页总字数控制在 150~320 字（信息充实但结构清晰，用换行分点）。
3. icon 字段只能从以下选：${VALID_ICONS.join(', ')}。

【JSON 输出格式示例】
\`\`\`json
{
  "projectName": "提炼的项目/汇报名称",
  "slides": [
    { "layout": "cover", "title": "2026年度工作述职与经营汇报", "subtitle": "汇报人：XXX · 日期" },
    {
      "layout": "dashboard",
      "title": "年度经营总览：核心指标实现双位数高质量增长",
      "subtitle": "全年签约与净利润均超额达成目标",
      "kpis": [
        { "label": "全年签约额", "value": "1023亿", "trend": "同比 +30.0%" },
        { "label": "创造净利润", "value": "141.4亿", "trend": "同比 +37.4%" },
        { "label": "滚动回款率", "value": "96.3%", "trend": "提升 +2.5pt" }
      ],
      "chartType": "groupedBar",
      "chartSeriesName": "2025年",
      "chartSeriesName2": "2024年",
      "chartUnit": "单位：亿元",
      "chartData": [
        { "label": "Q1", "value": 210, "value2": 165 },
        { "label": "Q2", "value": 255, "value2": 190 },
        { "label": "Q3", "value": 268, "value2": 205 },
        { "label": "Q4", "value": 290, "value2": 227 }
      ],
      "items": [
        { "head": "增长引擎", "body": "Q4单季签约突破 **290亿**，较去年同期增长 **27.8%**。" },
        { "head": "利润结构", "body": "高毛利产品占比提升至 **42%**，带动净利增速高于营收。" }
      ]
    },
    {
      "layout": "table",
      "title": "各板块业绩明细：核心业务全面跑赢预算",
      "subtitle": "重点指标同比均保持正增长",
      "tableData": {
        "headers": ["业务板块", "2024实际", "2025实际", "同比增长"],
        "rows": [
          ["华东大区", "320亿", "415亿", "+29.7%"],
          ["华南大区", "260亿", "348亿", "+33.8%"]
        ]
      },
      "items": [
        { "head": "区域亮点", "body": "华南大区增速达 **+33.8%**，位居全国首位。" },
        { "head": "改进方向", "body": "需进一步提升华北区新客转化率至 **25%** 以上。" }
      ]
    },
    {
      "layout": "fourCol",
      "title": "节点营销复盘：四大战役稳固全年业绩",
      "subtitle": "围绕大促节点形成标准化打法",
      "items": [
        { "head": "Q1 新春大促", "body": "• 聚焦会员专享复购\\n• 会员消费占比达 **65%**", "badge": "GMV 3.25亿" },
        { "head": "Q2 618大促", "body": "• 跨界联名新品首发\\n• 社媒总曝光超 **2亿**", "badge": "GMV 3.85亿" },
        { "head": "Q3 双11预热", "body": "• 全品类直播矩阵带货\\n• 新增转化率达 **30%**", "badge": "GMV 3.20亿" },
        { "head": "Q4 年终清仓", "body": "• 老客裂变与礼盒清仓\\n• 周转率提升 **40%**", "badge": "GMV 2.10亿" }
      ]
    }
  ]
}
\`\`\`

【以下是客户原始文档内容】
${rawText.slice(0, 12000)}`;

    try {
      await navigator.clipboard.writeText(prompt);
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    } catch {
      alert('复制失败，请检查浏览器剪贴板权限。');
    }
  };

  const handleParseAndGenerate = () => {
    if (!jsonInput.trim()) return alert('请先粘贴 AI 返回的 JSON 内容！');
    try {
      let cleaned = jsonInput.trim();
      const fenceMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/i);
      if (fenceMatch) cleaned = fenceMatch[1].trim();

      const parsed = JSON.parse(cleaned);
      const rawSlides = Array.isArray(parsed) ? parsed : parsed.slides;
      if (!Array.isArray(rawSlides) || rawSlides.length === 0) throw new Error('未找到有效的 slides 数组');

      const normalizedSlides = rawSlides.map((s, idx) => {
        const layout = VALID_LAYOUTS.includes(s.layout) ? s.layout : 'cards';
        const items = Array.isArray(s.items)
          ? s.items.map((it, i) => ({
              head: String(it.head || ''),
              body: String(it.body || ''),
              badge: String(it.badge || ''),
              icon: VALID_ICONS.includes(it.icon) ? it.icon : VALID_ICONS[i % VALID_ICONS.length],
            }))
          : [];
        while (items.length < 4) {
          items.push({ head: '', body: '', badge: '', icon: VALID_ICONS[items.length % VALID_ICONS.length] });
        }

        const chartData =
          Array.isArray(s.chartData) && s.chartData.length >= 2
            ? s.chartData.slice(0, 6).map((d) => ({
                label: String(d.label || '项'),
                value: Number(d.value) || 0,
                value2: d.value2 !== undefined ? Number(d.value2) || 0 : 0,
              }))
            : [
                { label: 'Q1', value: 120, value2: 95 },
                { label: 'Q2', value: 185, value2: 140 },
                { label: 'Q3', value: 240, value2: 190 },
                { label: 'Q4', value: 320, value2: 245 },
              ];

        const kpis =
          Array.isArray(s.kpis) && s.kpis.length >= 1
            ? s.kpis.slice(0, 3).map((k) => ({
                label: String(k.label || '核心指标'),
                value: String(k.value || '0'),
                trend: String(k.trend || ''),
              }))
            : [
                { label: '核心指标一', value: '1023亿', trend: '同比 +30%' },
                { label: '核心指标二', value: '141.4亿', trend: '同比 +37%' },
                { label: '核心指标三', value: '96.3%', trend: '提升 +2.5pt' },
              ];

        const tableData =
          s.tableData && Array.isArray(s.tableData.headers) && Array.isArray(s.tableData.rows)
            ? {
                headers: s.tableData.headers.map(String),
                rows: s.tableData.rows.map((r) => (Array.isArray(r) ? r.map(String) : [])),
              }
            : {
                headers: ['核心指标', '去年实际', '今年实际', '同比'],
                rows: [
                  ['签约额', '787亿', '1023亿', '+30%'],
                  ['净利润', '102亿', '141亿', '+37%'],
                ],
              };

        return {
          id: Date.now() + idx,
          layout,
          title: String(s.title || '未命名标题'),
          subtitle: String(s.subtitle || ''),
          image: '',
          kpis,
          chartType: VALID_CHARTS.includes(s.chartType) ? s.chartType : 'groupedBar',
          chartSeriesName: String(s.chartSeriesName || '2025年'),
          chartSeriesName2: s.chartSeriesName2 !== undefined ? String(s.chartSeriesName2) : '2024年',
          chartUnit: String(s.chartUnit || ''),
          chartData,
          tableData,
          items,
        };
      });

      onGenerate(normalizedSlides, parsed.projectName || fileName || 'AI生成汇报方案');
    } catch (err) {
      alert(`JSON 解析失败，请确认粘贴的是完整的 AI 输出内容。\n错误详情：${err.message}`);
    }
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(2, 6, 23, 0.8)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 20 }}>
      <div style={{ width: '100%', maxWidth: 820, maxHeight: '90vh', background: '#1e293b', border: '1px solid #38bdf8', borderRadius: 12, padding: 22, display: 'flex', flexDirection: 'column', gap: 14, overflowY: 'auto', color: '#f8fafc' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 'bold', color: '#38bdf8' }}>✨ 文档自动生成高密汇报 PPT</div>
            <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 2 }}>支持自动拆解为：经营仪表盘、原生表格、四栏季度页、多系列图表及重点标红</div>
          </div>
          <button onClick={onClose} style={{ background: '#334155', color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px', cursor: 'pointer' }}>关闭</button>
        </div>

        <div style={{ padding: 12, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>第 1 步：导入客户文档（支持 .docx / .txt / .md）</span>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <label style={{ fontSize: 12, color: '#94a3b8' }}>
                目标页数：
                <input type="number" min={4} max={25} value={slideCount} onChange={(e) => setSlideCount(Number(e.target.value) || 10)} style={{ width: 48, marginLeft: 4, padding: 3, background: '#1e293b', color: '#38bdf8', border: '1px solid #475569', borderRadius: 4, textAlign: 'center' }} />
              </label>
              <button onClick={() => fileInputRef.current.click()} style={{ padding: '6px 12px', fontSize: 12, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer' }}>
                + 上传 Word / TXT
              </button>
              <input ref={fileInputRef} type="file" accept=".docx,.txt,.md" onChange={handleFileUpload} style={{ display: 'none' }} />
            </div>
          </div>
          <textarea rows={5} value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="上传 Word (.docx) / TXT 文件，或直接把客户发来的长文案粘贴到这里..." style={{ width: '100%', padding: 8, background: '#1e293b', color: '#f8fafc', border: '1px solid #334155', borderRadius: 6, fontSize: 12, boxSizing: 'border-box', resize: 'vertical' }} />
          <button onClick={handleCopyPrompt} style={{ padding: 10, background: copied ? '#059669' : '#2563eb', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 13, fontWeight: 'bold' }}>
            {copied ? '✅ 已复制高密汇报提示词！去发给 AI 后把 JSON 贴在下方' : '📋 第 2 步：一键复制专属 AI 提示词（含文档内容）'}
          </button>
        </div>

        <div style={{ padding: 12, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>第 3 步：粘贴 AI 返回的 JSON，一键生成全套 PPT</div>
          <textarea rows={6} value={jsonInput} onChange={(e) => setJsonInput(e.target.value)} placeholder='把 AI 回复的整段内容（包含 ```json ... ``` 也没关系）直接粘贴到这里...' style={{ width: '100%', padding: 8, background: '#1e293b', color: '#4ade80', border: '1px solid #334155', borderRadius: 6, fontSize: 12, fontFamily: 'monospace', boxSizing: 'border-box', resize: 'vertical' }} />
          <button onClick={handleParseAndGenerate} style={{ padding: 12, background: '#059669', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 14, fontWeight: 'bold' }}>
            🚀 一键生成全套幻灯片（随后手动精修）
          </button>
        </div>
      </div>
    </div>
  );
}