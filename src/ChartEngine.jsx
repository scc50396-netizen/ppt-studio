export const CHART_TYPES = [
  { id: 'groupedBar', name: '对比柱状图（双系列）' },
  { id: 'line', name: '趋势折线图（单/双系列）' },
  { id: 'doughnut', name: '占比环形图' },
  { id: 'bar', name: '单系列柱状图' },
  { id: 'pie', name: '实心饼图' },
];

export const DEFAULT_CHART_DATA = [
  { label: 'Q1', value: 120, value2: 95 },
  { label: 'Q2', value: 185, value2: 140 },
  { label: 'Q3', value: 240, value2: 190 },
  { label: 'Q4', value: 320, value2: 245 },
];

export const isMultiSeries = (type) => type === 'groupedBar' || type === 'line';

const hex = (color) => (color || '#000000').replace('#', '');
const getPalette = (theme) => [theme.primary, theme.accent, '#10B981', '#8B5CF6', '#EC4899', '#F97316'];

// ================= 1. React SVG 实时预览组件 =================
export function ChartPreview({ type = 'groupedBar', data, theme, series1 = '2025年', series2 = '2024年', unit = '' }) {
  const safeData = data && data.length > 0 ? data : DEFAULT_CHART_DATA;
  const palette = getPalette(theme);
  const showS2 = isMultiSeries(type) && Boolean(series2 && series2.trim());

  // 环形图 / 实心饼图
  if (type === 'pie' || type === 'doughnut') {
    const total = safeData.reduce((sum, d) => sum + Math.max(0, Number(d.value) || 0), 0) || 1;
    let currentAngle = -Math.PI / 2;
    const cx = 128;
    const cy = 115;
    const r = 88;

    const slices = safeData.map((d, idx) => {
      const val = Math.max(0, Number(d.value) || 0);
      const ratio = val / total;
      const angle = ratio * Math.PI * 2;
      const startAngle = currentAngle;
      const endAngle = currentAngle + angle;
      currentAngle = endAngle;
      const x1 = cx + r * Math.cos(startAngle);
      const y1 = cy + r * Math.sin(startAngle);
      const x2 = cx + r * Math.cos(endAngle);
      const y2 = cy + r * Math.sin(endAngle);
      const largeArc = angle > Math.PI ? 1 : 0;
      const pathData = ratio >= 0.9999 ? '' : `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
      return { ...d, val, pct: Math.round(ratio * 100), color: palette[idx % palette.length], pathData, isFull: ratio >= 0.9999 };
    });

    return (
      <div style={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box' }}>
        {unit && <div style={{ fontSize: 10, color: theme.subText, paddingLeft: 4 }}>{unit}</div>}
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'space-around', minHeight: 0 }}>
          <svg viewBox="0 0 256 230" style={{ width: '60%', height: '100%', maxHeight: 225 }}>
            {slices.map((s, idx) =>
              s.isFull ? <circle key={idx} cx={cx} cy={cy} r={r} fill={s.color} /> : <path key={idx} d={s.pathData} fill={s.color} stroke={theme.cardBg} strokeWidth="2" />
            )}
            {type === 'doughnut' && (
              <>
                <circle cx={cx} cy={cy} r={52} fill={theme.cardBg} />
                <text x={cx} y={cy - 4} textAnchor="middle" fill={theme.text} fontSize="18" fontWeight="900">
                  {Math.round(total * 10) / 10}
                </text>
                <text x={cx} y={cy + 15} textAnchor="middle" fill={theme.subText} fontSize="10">
                  合计总览
                </text>
              </>
            )}
          </svg>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 7, width: '38%' }}>
            {slices.map((s, idx) => (
              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11.5 }}>
                <span style={{ width: 9, height: 9, borderRadius: 2, background: s.color, flexShrink: 0 }} />
                <span style={{ color: theme.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }}>{s.label}</span>
                <span style={{ color: theme.subText, fontWeight: 'bold' }}>{s.pct}%</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // 柱状图 / 对比柱状图 / 折线图
  const allVals = safeData.flatMap((d) => (showS2 ? [Math.max(0, Number(d.value) || 0), Math.max(0, Number(d.value2) || 0)] : [Math.max(0, Number(d.value) || 0)]));
  const maxVal = Math.max(...allVals, 1);
  const chartW = 440;
  const chartH = 225;
  const padL = 38;
  const padR = 16;
  const padT = 36;
  const padB = 34;
  const innerW = chartW - padL - padR;
  const innerH = chartH - padT - padB;
  const slotW = innerW / safeData.length;

  return (
    <svg viewBox={`0 0 ${chartW} ${chartH}`} style={{ width: '100%', height: '100%' }}>
      {/* 左上角单位 & 右上角图例 */}
      {unit && (
        <text x={padL} y={14} fill={theme.subText} fontSize="10">
          {unit}
        </text>
      )}
      {isMultiSeries(type) && (
        <g>
          <rect x={chartW - (showS2 ? 155 : 80)} y={5} width={10} height={10} rx="2" fill={theme.primary} />
          <text x={chartW - (showS2 ? 141 : 66)} y={14} fill={theme.text} fontSize="10">{(series1 || '系列1').slice(0, 6)}</text>
          {showS2 && (
            <>
              <rect x={chartW - 78} y={5} width={10} height={10} rx="2" fill={theme.accent} />
              <text x={chartW - 64} y={14} fill={theme.text} fontSize="10">{(series2 || '系列2').slice(0, 6)}</text>
            </>
          )}
        </g>
      )}

      {/* 网格线 */}
      {[0, 0.5, 1].map((t, i) => {
        const y = padT + innerH * (1 - t);
        return (
          <g key={i}>
            <line x1={padL} y1={y} x2={chartW - padR} y2={y} stroke={theme.subText} strokeOpacity="0.22" strokeDasharray="3 3" />
            <text x={padL - 6} y={y + 4} textAnchor="end" fill={theme.subText} fontSize="10">{Math.round(maxVal * t)}</text>
          </g>
        );
      })}

      {/* 折线图 */}
      {type === 'line' && (() => {
        const pts1 = safeData.map((d, idx) => ({
          x: padL + idx * slotW + slotW / 2,
          y: padT + innerH - (Math.max(0, Number(d.value) || 0) / maxVal) * innerH,
          val: Math.max(0, Number(d.value) || 0),
          label: d.label,
        }));
        const pts2 = showS2
          ? safeData.map((d, idx) => ({
              x: padL + idx * slotW + slotW / 2,
              y: padT + innerH - (Math.max(0, Number(d.value2) || 0) / maxVal) * innerH,
              val: Math.max(0, Number(d.value2) || 0),
            }))
          : [];
        return (
          <g>
            {showS2 && (
              <polyline fill="none" stroke={theme.accent} strokeWidth="2.5" points={pts2.map((p) => `${p.x},${p.y}`).join(' ')} />
            )}
            <polyline fill="none" stroke={theme.primary} strokeWidth="3" points={pts1.map((p) => `${p.x},${p.y}`).join(' ')} />
            {pts2.map((p, i) => (
              <g key={`p2-${i}`}>
                <circle cx={p.x} cy={p.y} r="4" fill={theme.cardBg} stroke={theme.accent} strokeWidth="2" />
                <text x={p.x} y={p.y + 14} textAnchor="middle" fill={theme.accent} fontSize="10" fontWeight="bold">{p.val}</text>
              </g>
            ))}
            {pts1.map((p, i) => (
              <g key={`p1-${i}`}>
                <circle cx={p.x} cy={p.y} r="4.5" fill={theme.cardBg} stroke={theme.primary} strokeWidth="2.5" />
                <text x={p.x} y={p.y - 8} textAnchor="middle" fill={theme.text} fontSize="10.5" fontWeight="bold">{p.val}</text>
                <text x={p.x} y={chartH - 10} textAnchor="middle" fill={theme.subText} fontSize="10.5">{(p.label || '').slice(0, 6)}</text>
              </g>
            ))}
          </g>
        );
      })()}

      {/* 对比柱状图（双系列） */}
      {type === 'groupedBar' && safeData.map((d, idx) => {
        const val1 = Math.max(0, Number(d.value) || 0);
        const val2 = Math.max(0, Number(d.value2) || 0);
        const h1 = (val1 / maxVal) * innerH;
        const h2 = (val2 / maxVal) * innerH;
        const barW = Math.min(26, slotW * 0.34);
        const centerX = padL + idx * slotW + slotW / 2;
        const x1 = showS2 ? centerX - barW - 2 : centerX - barW / 2;
        const x2 = centerX + 2;
        return (
          <g key={idx}>
            <rect x={x1} y={padT + innerH - h1} width={barW} height={Math.max(2, h1)} rx="3" fill={theme.primary} />
            <text x={x1 + barW / 2} y={padT + innerH - h1 - 5} textAnchor="middle" fill={theme.text} fontSize="10" fontWeight="bold">{val1}</text>
            {showS2 && (
              <>
                <rect x={x2} y={padT + innerH - h2} width={barW} height={Math.max(2, h2)} rx="3" fill={theme.accent} />
                <text x={x2 + barW / 2} y={padT + innerH - h2 - 5} textAnchor="middle" fill={theme.accent} fontSize="10" fontWeight="bold">{val2}</text>
              </>
            )}
            <text x={centerX} y={chartH - 10} textAnchor="middle" fill={theme.subText} fontSize="10.5">{(d.label || '').slice(0, 6)}</text>
          </g>
        );
      })}

      {/* 单系列柱状图 */}
      {type === 'bar' && safeData.map((d, idx) => {
        const val = Math.max(0, Number(d.value) || 0);
        const h = (val / maxVal) * innerH;
        const barW = Math.min(44, slotW * 0.56);
        const x = padL + idx * slotW + (slotW - barW) / 2;
        const y = padT + innerH - h;
        return (
          <g key={idx}>
            <rect x={x} y={y} width={barW} height={Math.max(2, h)} rx="4" fill={idx === safeData.length - 1 ? theme.accent : theme.primary} />
            <text x={x + barW / 2} y={y - 6} textAnchor="middle" fill={theme.text} fontSize="11" fontWeight="bold">{val}</text>
            <text x={x + barW / 2} y={chartH - 10} textAnchor="middle" fill={theme.subText} fontSize="10.5">{(d.label || '').slice(0, 6)}</text>
          </g>
        );
      })}
    </svg>
  );
}

// ================= 2. PptxGenJS 原生可编辑图表写入器 =================
export function addChartToPptxSlide(pptx, slide, { x, y, w, h, type = 'groupedBar', data, theme, series1 = '2025年', series2 = '2024年', unit = '' }) {
  const safeData = data && data.length > 0 ? data : DEFAULT_CHART_DATA;
  const paletteHex = getPalette(theme).map(hex);
  const showS2 = isMultiSeries(type) && Boolean(series2 && series2.trim());

  const labels = safeData.map((d) => d.label || '项');
  const values1 = safeData.map((d) => Math.max(0, Number(d.value) || 0));
  const values2 = safeData.map((d) => Math.max(0, Number(d.value2) || 0));

  const chartSeries = showS2
    ? [
        { name: series1 || '系列1', labels, values: values1 },
        { name: series2 || '系列2', labels, values: values2 },
      ]
    : [{ name: series1 || '数据指标', labels, values: values1 }];

  if (unit) {
    slide.addText(unit, { x, y: y - 0.22, w: 2.5, h: 0.24, fontSize: 9.5, color: hex(theme.subText) });
  }

  if (type === 'pie') {
    slide.addChart(pptx.ChartType.pie, [chartSeries[0]], {
      x, y, w, h,
      chartColors: paletteHex,
      showPercent: true,
      showLegend: true,
      legendPos: 'r',
      legendColor: hex(theme.text),
      
    });
  } else if (type === 'doughnut') {
    slide.addChart(pptx.ChartType.doughnut, [chartSeries[0]], {
      x, y, w, h,
      holeSize: 55,
      chartColors: paletteHex,
      showPercent: true,
      showLegend: true,
      legendPos: 'r',
      legendColor: hex(theme.text),
      
    });
  } else if (type === 'line') {
    slide.addChart(pptx.ChartType.line, chartSeries, {
      x, y, w, h,
      chartColors: [hex(theme.primary), hex(theme.accent)],
      lineSize: 3,
      lineDataSymbol: 'circle',
      showValue: true,
      
      catAxisLabelColor: hex(theme.subText),
      valAxisLabelColor: hex(theme.subText),
      showLegend: showS2,
      legendPos: 't',
      legendColor: hex(theme.text),
    });
  } else {
    // groupedBar 或 bar
    slide.addChart(pptx.ChartType.bar, chartSeries, {
      x, y, w, h,
      barDir: 'col',
      barGrouping: 'clustered',
      chartColors: showS2 ? [hex(theme.primary), hex(theme.accent)] : [hex(theme.primary)],
      showValue: true,
      
      catAxisLabelColor: hex(theme.subText),
      valAxisLabelColor: hex(theme.subText),
      showLegend: showS2,
      legendPos: 't',
      legendColor: hex(theme.text),
    });
  }
}

// ================= 3. Canvas 图表绘制器（用于微信长图与水印图） =================
export function drawChartOnCanvas(ctx, { x, y, w, h, type = 'groupedBar', data, theme, series1 = '2025年', series2 = '2024年', unit = '', fontFamily = 'sans-serif' }) {
  const safeData = data && data.length > 0 ? data : DEFAULT_CHART_DATA;
  const palette = getPalette(theme);
  const showS2 = isMultiSeries(type) && Boolean(series2 && series2.trim());

  ctx.fillStyle = theme.cardBg;
  ctx.fillRect(x, y, w, h);

  if (unit) {
    ctx.fillStyle = theme.subText;
    ctx.font = `14px "${fontFamily}", sans-serif`;
    ctx.textAlign = 'left';
    ctx.fillText(unit, x + 20, y + 26);
  }

  if (type === 'pie' || type === 'doughnut') {
    const total = safeData.reduce((s, d) => s + Math.max(0, Number(d.value) || 0), 0) || 1;
    let angle = -Math.PI / 2;
    const cx = x + w * 0.34;
    const cy = y + h * 0.54;
    const r = Math.min(w * 0.26, h * 0.36);

    safeData.forEach((d, idx) => {
      const val = Math.max(0, Number(d.value) || 0);
      const sliceAngle = (val / total) * Math.PI * 2;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, angle, angle + sliceAngle);
      ctx.closePath();
      ctx.fillStyle = palette[idx % palette.length];
      ctx.fill();
      angle += sliceAngle;

      const ly = y + 60 + idx * 36;
      ctx.fillStyle = palette[idx % palette.length];
      ctx.fillRect(x + w * 0.65, ly - 13, 14, 14);
      ctx.fillStyle = theme.text;
      ctx.font = `16px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'left';
      ctx.fillText(`${d.label} (${Math.round((val / total) * 100)}%)`, x + w * 0.65 + 22, ly);
    });

    if (type === 'doughnut') {
      ctx.beginPath();
      ctx.arc(cx, cy, r * 0.58, 0, Math.PI * 2);
      ctx.fillStyle = theme.cardBg;
      ctx.fill();
      ctx.fillStyle = theme.text;
      ctx.font = `bold 24px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'center';
      ctx.fillText(String(Math.round(total * 10) / 10), cx, cy + 6);
    }
    ctx.textAlign = 'left';
    return;
  }

  const allVals = safeData.flatMap((d) => (showS2 ? [Math.max(0, Number(d.value) || 0), Math.max(0, Number(d.value2) || 0)] : [Math.max(0, Number(d.value) || 0)]));
  const maxVal = Math.max(...allVals, 1);
  const padL = 42;
  const padR = 24;
  const padT = 56;
  const padB = 46;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;
  const baseY = y + padT + plotH;
  const slotW = plotW / safeData.length;

  // 图例
  if (isMultiSeries(type)) {
    ctx.fillStyle = theme.primary;
    ctx.fillRect(x + w - 210, y + 14, 14, 14);
    ctx.fillStyle = theme.text;
    ctx.font = `14px "${fontFamily}", sans-serif`;
    ctx.textAlign = 'left';
    ctx.fillText((series1 || '系列1').slice(0, 6), x + w - 190, y + 26);
    if (showS2) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(x + w - 105, y + 14, 14, 14);
      ctx.fillStyle = theme.text;
      ctx.fillText((series2 || '系列2').slice(0, 6), x + w - 85, y + 26);
    }
  }

  if (type === 'line') {
    const drawLine = (vals, color, offsetY) => {
      ctx.beginPath();
      ctx.strokeStyle = color;
      ctx.lineWidth = 4;
      vals.forEach((v, idx) => {
        const px = x + padL + idx * slotW + slotW / 2;
        const py = baseY - (v / maxVal) * plotH;
        if (idx === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
      });
      ctx.stroke();

      vals.forEach((v, idx) => {
        const px = x + padL + idx * slotW + slotW / 2;
        const py = baseY - (v / maxVal) * plotH;
        ctx.beginPath();
        ctx.arc(px, py, 6, 0, Math.PI * 2);
        ctx.fillStyle = color;
        ctx.fill();
        ctx.fillStyle = color;
        ctx.font = `bold 15px "${fontFamily}", sans-serif`;
        ctx.textAlign = 'center';
        ctx.fillText(String(v), px, py + offsetY);
      });
    };
    if (showS2) drawLine(safeData.map((d) => Math.max(0, Number(d.value2) || 0)), theme.accent, 22);
    drawLine(safeData.map((d) => Math.max(0, Number(d.value) || 0)), theme.primary, -12);
    safeData.forEach((d, idx) => {
      ctx.fillStyle = theme.subText;
      ctx.font = `15px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'center';
      ctx.fillText((d.label || '').slice(0, 6), x + padL + idx * slotW + slotW / 2, baseY + 26);
    });
  } else {
    safeData.forEach((d, idx) => {
      const val1 = Math.max(0, Number(d.value) || 0);
      const val2 = Math.max(0, Number(d.value2) || 0);
      const centerX = x + padL + idx * slotW + slotW / 2;

      if (type === 'groupedBar' && showS2) {
        const barW = Math.min(38, slotW * 0.34);
        const h1 = (val1 / maxVal) * plotH;
        const h2 = (val2 / maxVal) * plotH;
        ctx.fillStyle = theme.primary;
        ctx.fillRect(centerX - barW - 3, baseY - h1, barW, Math.max(4, h1));
        ctx.fillStyle = theme.accent;
        ctx.fillRect(centerX + 3, baseY - h2, barW, Math.max(4, h2));
        ctx.textAlign = 'center';
        ctx.font = `bold 15px "${fontFamily}", sans-serif`;
        ctx.fillStyle = theme.text;
        ctx.fillText(String(val1), centerX - barW / 2 - 3, baseY - h1 - 8);
        ctx.fillStyle = theme.accent;
        ctx.fillText(String(val2), centerX + barW / 2 + 3, baseY - h2 - 8);
      } else {
        const barW = Math.min(60, slotW * 0.54);
        const h1 = (val1 / maxVal) * plotH;
        ctx.fillStyle = idx === safeData.length - 1 ? theme.accent : theme.primary;
        ctx.fillRect(centerX - barW / 2, baseY - h1, barW, Math.max(4, h1));
        ctx.textAlign = 'center';
        ctx.fillStyle = theme.text;
        ctx.font = `bold 16px "${fontFamily}", sans-serif`;
        ctx.fillText(String(val1), centerX, baseY - h1 - 8);
      }
      ctx.fillStyle = theme.subText;
      ctx.font = `15px "${fontFamily}", sans-serif`;
      ctx.textAlign = 'center';
      ctx.fillText((d.label || '').slice(0, 6), centerX, baseY + 26);
    });
  }
  ctx.textAlign = 'left';
}