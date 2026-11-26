// 富文本解析：将 **关键词** 渲染为 React 加粗标红元素
export function renderRichText(text, accentColor) {
  if (!text) return null;
  const parts = String(text).split(/(\*\*.*?\*\*)/g);
  return parts.map((part, idx) => {
    if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
      return (
        <strong key={idx} style={{ color: accentColor, fontWeight: 900 }}>
          {part.slice(2, -2)}
        </strong>
      );
    }
    return <span key={idx}>{part}</span>;
  });
}

// 富文本转换：将 **关键词** 转为 PptxGenJS 原生可编辑富文本数组
export function toPptxRichText(text, baseOpts, accentHex) {
  const str = String(text || '');
  const parts = str.split(/(\*\*.*?\*\*)/g).filter(Boolean);
  if (parts.length === 0) return [{ text: '', options: baseOpts }];
  return parts.map((part) => {
    if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
      return {
        text: part.slice(2, -2),
        options: { ...baseOpts, bold: true, color: accentHex },
      };
    }
    return { text: part, options: baseOpts };
  });
}

// 去除 ** 符号（用于 Canvas 长图绘制和字数统计）
export const stripMarkdown = (text) => String(text || '').replace(/\*\*(.*?)\*\*/g, '$1');

export const DEFAULT_TABLE_DATA = {
  headers: ['核心指标', '2024年实际', '2025年目标', '同比增长'],
  rows: [
    ['区域签约额(亿元)', '787.0', '1023.0', '+30.0%'],
    ['创造净利润(亿元)', '102.9', '141.4', '+37.4%'],
    ['结算净利润(亿元)', '63.4', '111.4', '+75.7%'],
    ['滚动12月回款率', '93.8%', '96.3%', '+2.5pt'],
  ],
};