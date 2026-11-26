import { useState, useRef } from 'react';
import mammoth from 'mammoth';

const VALID_LAYOUTS = ['cover', 'toc', 'section', 'dashboard', 'table', 'fourCol', 'chart', 'compare', 'imageText', 'cards', 'timeline', 'metrics', 'ending'];
const VALID_ICONS = ['trend', 'rocket', 'shield', 'target', 'zap', 'award', 'users', 'bulb', 'layers', 'clock', 'chart', 'globe'];
const VALID_CHARTS = ['groupedBar', 'line', 'doughnut', 'bar', 'pie'];

// 4 套高频接单提示词风格引擎
const PROMPT_STYLES = [
  {
    id: 'report',
    name: '📊 大厂/咨询高密汇报',
    desc: '150~320字/页 · 结论先行 · 表格/仪表盘密集',
    role: '顶级大厂/咨询风商业 PPT 策划师（高密阅读型工作汇报风格）',
    rules: `1. 必须包含封面(cover)、目录(toc)、结尾(ending)，中间页优先使用高密数据版式：dashboard（经营仪表盘）、table（商务数据表格）、fourCol（四栏/四季度）、chart（多系列对比图）。
2. 每页主标题采用“模块名：一句话核心结论”格式；正文用换行分点，单页总字数控制在 150~320 字，信息充实有业务颗粒度。
3. 必须用 **关键词** 包裹核心数字、同比增速或关键结论（系统会自动加粗标红）。
4. 每页在 notes 字段附带 60~100 字的精炼汇报口播提示。`,
  },
  {
    id: 'thesis',
    name: '🎓 毕业论文/学术答辩',
    desc: '学术严谨 · 逻辑闭环 · 自动生成答辩逐字稿',
    role: '高校本科/硕士毕业论文答辩 PPT 专家',
    rules: `1. 严格遵循学术答辩逻辑拆解：封面(cover) → 目录(toc) → 研究背景与意义 → 国内外现状/痛点 → 研究方法与框架 → 核心模型/实验数据(优先用 table/dashboard/chart 呈现量化对比) → 结论与展望 → 致谢(ending)。
2. 语言客观严谨、术语准确，提炼论文核心论点与实测数据，避免大段照抄原文，单页总字数控制在 140~260 字。
3. 用 **关键词** 标红核心实验指标、创新点或提升幅度。
4. 必须在每页的 notes 字段撰写 80~150 字的自然流畅【答辩口播逐字稿】，方便学生直接照着念。`,
  },
  {
    id: 'pitch',
    name: '🚀 商业路演/融资推介',
    desc: '60~120字/页 · 大字报冲击力 · 强说服力',
    role: '顶级投资机构路演 BP 与发布会 PPT 策划师',
    rules: `1. 按商业路演黄金结构拆解：封面(cover) → 市场痛点(compare) → 核心解决方案(cards/imageText) → 市场规模与增长(chart/metrics) → 商业模式与壁垒(fourCol) → 里程碑规划(timeline) → 融资/合作结尾(ending)。
2. 遵循“演讲极简原则”：单页总字数严格控制在 60~120 字，短句金句化，留白充足，绝不堆砌长段落。
3. 优先使用 metrics（核心大数字）、compare（痛点vs方案）、cards（三栏图标）突出视觉冲击力，并用 **关键词** 标红核心变现数据与优势。
4. 在每页 notes 字段写入 80~120 字的感染力路演演讲词。`,
  },
  {
    id: 'gov',
    name: '🏛️ 政务/党建/国企汇报',
    desc: '措辞庄重 · 排比工整 · 突出成效与规划',
    role: '资深政府机关及大型国企公文汇报 PPT 策划师',
    rules: `1. 结构稳重规范：封面(cover) → 目录(toc) → 总体成效回顾(dashboard/metrics) → 重点工作推进(fourCol/cards) → 量化指标完成情况(table/chart) → 存在不足与下步举措(compare/timeline) → 结尾(ending)。
2. 措辞严谨庄重、站位清晰，卡片小标题尽量采用字数对称的排比短语（如“强化统筹、聚焦发力、提质增效”），单页总字数 150~280 字。
3. 用 **关键词** 标红关键政绩数据、完成率与核心抓手。
4. 在每页 notes 字段撰写 60~100 字的正式汇报发言稿。`,
  },
];

export default function DocToPptModal({ onClose, onGenerate }) {
  const [rawText, setRawText] = useState('');
  const [fileName, setFileName] = useState('');
  const [slideCount, setSlideCount] = useState(10);
  const [styleId, setStyleId] = useState('report');
  const [customReq, setCustomReq] = useState('');
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const fileInputRef = useRef(null);

  const activeStyle = PROMPT_STYLES.find((s) => s.id === styleId) || PROMPT_STYLES[0];

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

    const extraLine = customReq.trim() ? `\n【客户额外定制要求（必须严格满足）】\n${customReq.trim()}\n` : '';

    const prompt = `你是一位${activeStyle.role}。请将以下文档拆解为一套约 ${slideCount} 页的 PPT，严格按下方 JSON 格式输出（只输出纯 JSON 代码块，不要多余解释）。

【当前选定风格：${activeStyle.name} · 排版与文案规范】
${activeStyle.rules}
5. 可用版式(layout)：cover, toc, section, dashboard, table, fourCol, chart, compare, imageText, cards, timeline, metrics, ending。
6. 图表类型(chartType)：groupedBar(双柱对比), line(折线), doughnut(环形), bar(单柱), pie(饼图)。
7. 图标(icon)只能从以下选：${VALID_ICONS.join(', ')}。
${extraLine}
【JSON 输出格式示例】
\`\`\`json
{
  "projectName": "提炼的项目/汇报名称",
  "slides": [
    {
      "layout": "cover",
      "title": "主标题",
      "subtitle": "副标题 · 汇报人 · 日期",
      "notes": "本页开场口播稿..."
    },
    {
      "layout": "dashboard",
      "title": "模块名：一句话核心结论",
      "subtitle": "补充说明",
      "notes": "本页数据解读口播稿...",
      "kpis": [
        { "label": "指标一", "value": "1023亿", "trend": "同比 +30.0%" },
        { "label": "指标二", "value": "141.4亿", "trend": "同比 +37.4%" },
        { "label": "指标三", "value": "96.3%", "trend": "提升 +2.5pt" }
      ],
      "chartType": "groupedBar",
      "chartSeriesName": "2025年",
      "chartSeriesName2": "2024年",
      "chartUnit": "单位：亿元",
      "chartData": [
        { "label": "Q1", "value": 210, "value2": 165 },
        { "label": "Q2", "value": 255, "value2": 190 }
      ],
      "items": [
        { "head": "归因一", "body": "核心数据突破 **290亿**，同比增长 **+27.8%**。" },
        { "head": "归因二", "body": "高毛利业务占比提升至 **42%**。" }
      ]
    },
    {
      "layout": "table",
      "title": "数据明细对比",
      "subtitle": "核心指标对比",
      "notes": "本页表格解读...",
      "tableData": {
        "headers": ["维度", "基准值", "优化值", "核心提升"],
        "rows": [
          ["指标A", "100", "135", "+35%"],
          ["指标B", "80%", "95%", "+15pt"]
        ]
      },
      "items": [
        { "head": "核心结论一", "body": "指标A大幅提升 **+35%**。" },
        { "head": "核心结论二", "body": "指标B达到 **95%** 行业领先水平。" }
      ]
    },
    {
      "layout": "fourCol",
      "title": "四大核心模块拆解",
      "subtitle": "结构化推进路径",
      "notes": "本页四栏解读...",
      "items": [
        { "head": "模块一", "body": "• 要点1\\n• 核心指标达 **65%**", "badge": "关键成果" },
        { "head": "模块二", "body": "• 要点1\\n• 核心指标达 **40%**", "badge": "关键成果" },
        { "head": "模块三", "body": "• 要点1\\n• 核心指标达 **30%**", "badge": "关键成果" },
        { "head": "模块四", "body": "• 要点1\\n• 核心指标达 **20%**", "badge": "关键成果" }
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
          notes: String(s.notes || ''),
          image: '',
          fontScale: 1,
          reverseLayout: false,
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

      onGenerate(normalizedSlides, parsed.projectName || fileName || 'AI生成演示方案');
    } catch (err) {
      alert(`JSON 解析失败，请确认粘贴的是完整的 AI 输出内容。\n错误详情：${err.message}`);
    }
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(2, 6, 23, 0.82)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 20 }}>
      <div style={{ width: '100%', maxWidth: 840, maxHeight: '92vh', background: '#1e293b', border: '1px solid #38bdf8', borderRadius: 12, padding: 22, display: 'flex', flexDirection: 'column', gap: 14, overflowY: 'auto', color: '#f8fafc' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 'bold', color: '#38bdf8' }}>✨ 文档一键生成 PPT（多风格提示词引擎）</div>
            <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 2 }}>按客户场景切换输出风格与字数密度，并自动同步生成每页演讲者逐字稿</div>
          </div>
          <button onClick={onClose} style={{ background: '#334155', color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px', cursor: 'pointer' }}>关闭</button>
        </div>

        {/* 新增：接单文本输出风格选择器 */}
        <div style={{ padding: 12, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>第 1 步：选择文本输出风格（决定字数密度、语气与版式倾向）</div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 8 }}>
            {PROMPT_STYLES.map((st) => {
              const active = styleId === st.id;
              return (
                <div
                  key={st.id}
                  onClick={() => setStyleId(st.id)}
                  style={{
                    padding: '9px 10px',
                    borderRadius: 6,
                    cursor: 'pointer',
                    background: active ? '#172554' : '#1e293b',
                    border: active ? '1.5px solid #38bdf8' : '1px solid #334155',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: 4,
                  }}
                >
                  <div style={{ fontSize: 12, fontWeight: 'bold', color: active ? '#38bdf8' : '#f8fafc' }}>{st.name}</div>
                  <div style={{ fontSize: 10.5, color: '#94a3b8', lineHeight: 1.4 }}>{st.desc}</div>
                </div>
              );
            })}
          </div>
          <input
            value={customReq}
            onChange={(e) => setCustomReq(e.target.value)}
            placeholder="选填：输入客户额外要求（如：重点突出降本增效、必须包含SWOT分析、语气简洁有力）"
            style={{ width: '100%', padding: '7px 10px', background: '#1e293b', color: '#fbbf24', border: '1px solid #475569', borderRadius: 6, fontSize: 12, boxSizing: 'border-box' }}
          />
        </div>

        <div style={{ padding: 12, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>第 2 步：导入客户文档（支持 .docx / .txt / .md）</span>
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
          <textarea rows={4} value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="上传 Word (.docx) / TXT 文件，或直接把客户发来的长文案粘贴到这里..." style={{ width: '100%', padding: 8, background: '#1e293b', color: '#f8fafc', border: '1px solid #334155', borderRadius: 6, fontSize: 12, boxSizing: 'border-box', resize: 'vertical' }} />
          <button onClick={handleCopyPrompt} style={{ padding: 10, background: copied ? '#059669' : '#2563eb', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 13, fontWeight: 'bold' }}>
            {copied ? `✅ 已复制【${activeStyle.name}】专属提示词！发给 AI 后把 JSON 贴在下方` : `📋 复制【${activeStyle.name}】专属 AI 提示词`}
          </button>
        </div>

        <div style={{ padding: 12, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>第 3 步：粘贴 AI 返回的 JSON，一键生成全套 PPT 与逐字稿</div>
          <textarea rows={5} value={jsonInput} onChange={(e) => setJsonInput(e.target.value)} placeholder='把 AI 回复的整段内容（包含 ```json ... ``` 也没关系）直接粘贴到这里...' style={{ width: '100%', padding: 8, background: '#1e293b', color: '#4ade80', border: '1px solid #334155', borderRadius: 6, fontSize: 12, fontFamily: 'monospace', boxSizing: 'border-box', resize: 'vertical' }} />
          <button onClick={handleParseAndGenerate} style={{ padding: 12, background: '#059669', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontSize: 14, fontWeight: 'bold' }}>
            🚀 一键生成全套幻灯片（含演讲者备注）
          </button>
        </div>
      </div>
    </div>
  );
}
