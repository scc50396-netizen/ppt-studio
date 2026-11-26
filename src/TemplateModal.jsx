 import { useState, useRef } from 'react';
import JSZip from 'jszip';

// 读取图片真实像素宽高
const getImageDimensions = (dataUrl) =>
  new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = dataUrl;
  });

// 分辨率清晰度评级
const getQualityTag = (w, h) => {
  if (w >= 1920 && h >= 1000) return { label: `高清 ${w}×${h}`, color: '#4ade80', bg: '#052e16' };
  if (w >= 1280 && h >= 700) return { label: `标清 ${w}×${h}`, color: '#38bdf8', bg: '#082f49' };
  return { label: `⚠️ 偏小易糊 ${w}×${h}`, color: '#fca5a5', bg: '#450a0a' };
};

export default function TemplateModal({ config, onClose, onApply }) {
  const [templateName, setTemplateName] = useState('我的专属商用模板');
  const [pptxInfo, setPptxInfo] = useState(null);
  const [showSmallImgs, setShowSmallImgs] = useState(false);
  const [parsing, setParsing] = useState(false);

  const coverInputRef = useRef(null);
  const contentInputRef = useRef(null);
  const importTplRef = useRef(null);
  const pptxInputRef = useRef(null);

  // 直接解析上传的参考 .pptx 模板（带真实分辨率检测与高清排序）
  const handlePptxUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setParsing(true);
    try {
      const zip = await JSZip.loadAsync(file);

      let primaryColor = '';
      let accentColor = '';
      let detectedFont = '';

      const themeFile = zip.file('ppt/theme/theme1.xml');
      if (themeFile) {
        const xmlText = await themeFile.async('text');
        const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
        const getColorFromTag = (tagName) => {
          const node = doc.getElementsByTagName(tagName)[0];
          if (!node) return '';
          const srgb = node.getElementsByTagName('a:srgbClr')[0];
          return srgb ? `#${srgb.getAttribute('val')}` : '';
        };
        primaryColor = getColorFromTag('a:accent1');
        accentColor = getColorFromTag('a:accent2');
        const eaFont = doc.getElementsByTagName('a:ea')[0];
        if (eaFont) detectedFont = eaFont.getAttribute('typeface') || '';
      }

      // 提取全部图片并读取真实宽高，按清晰度从高到低排序
      const mediaEntries = Object.keys(zip.files).filter((path) =>
        /^ppt\/media\/.*\.(png|jpe?g|webp)$/i.test(path)
      );

      const mediaImages = [];
      for (const path of mediaEntries.slice(0, 40)) {
        const ext = path.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg';
        const base64 = await zip.files[path].async('base64');
        const dataUrl = `data:${mime};base64,${base64}`;
        const { w, h } = await getImageDimensions(dataUrl);
        if (w > 0 && h > 0) {
          mediaImages.push({ name: path.split('/').pop(), dataUrl, w, h, pixels: w * h });
        }
      }

      // 按总像素降序排列（大图/高清底图永远排最前）
      mediaImages.sort((a, b) => b.pixels - a.pixels);

      const cleanName = file.name.replace(/\.pptx$/i, '');
      setTemplateName(cleanName);
      setPptxInfo({ fileName: cleanName, primaryColor, accentColor, detectedFont, images: mediaImages });

      if (primaryColor) {
        onApply(
          { customPrimary: primaryColor },
          `🎨 已提取模板「${cleanName}」主题主色 ${primaryColor}，已按清晰度排序展示图片！`
        );
      }
    } catch {
      alert('解析失败：请确认上传的是标准 .pptx 文件（暂不支持老版 .ppt 或加密文件）。');
    } finally {
      setParsing(false);
      e.target.value = '';
    }
  };

  // 手动上传封面/内页底图（带低分辨率防糊提醒）
  const handleBgUpload = (field) => (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      const dataUrl = ev.target.result;
      const { w, h } = await getImageDimensions(dataUrl);
      const isLowRes = w < 1280 || h < 700;
      const msg = isLowRes
        ? `⚠️ 该底图分辨率偏低（${w}×${h}），放大后可能模糊！建议使用 1920×1080 或更高清的 16:9 底图。`
        : `✅ 已应用高清底图（${w}×${h}），请确认素材无版权风险。`;
      onApply({ [field]: dataUrl }, msg);
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const handleExportTemplate = () => {
    const payload = {
      type: 'ppt-studio-template',
      version: 1,
      templateName: templateName || '自定义模板',
      exportedAt: new Date().toISOString(),
      themeId: config.themeId,
      customPrimary: config.customPrimary,
      fontFamily: config.fontFamily,
      brandLogo: config.brandLogo,
      showFooter: config.showFooter,
      coverBg: config.coverBg,
      contentBg: config.contentBg,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${templateName || '商用风格'}-PPT模板包.json`;
    link.click();
  };

  const handleImportTemplate = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!data.themeId && !data.coverBg && !data.contentBg) throw new Error('无效模板');
        onApply(
          {
            themeId: data.themeId || config.themeId,
            customPrimary: data.customPrimary ?? '',
            fontFamily: data.fontFamily || config.fontFamily,
            brandLogo: data.brandLogo ?? '',
            showFooter: typeof data.showFooter === 'boolean' ? data.showFooter : config.showFooter,
            coverBg: data.coverBg || '',
            contentBg: data.contentBg || '',
          },
          `🎨 已套用模板「${data.templateName || file.name}」，原有文案已完整保留！`
        );
        onClose();
      } catch {
        alert('导入失败：请选择本软件导出的「PPT模板包.json」文件。');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const visibleImages = pptxInfo
    ? pptxInfo.images.filter((img) => showSmallImgs || (img.w >= 600 && img.h >= 340))
    : [];
  const hiddenCount = pptxInfo ? pptxInfo.images.length - visibleImages.length : 0;

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(2, 6, 23, 0.8)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 20 }}>
      <div style={{ width: '100%', maxWidth: 740, maxHeight: '90vh', overflowY: 'auto', background: '#1e293b', border: '1px solid #38bdf8', borderRadius: 12, padding: 22, display: 'flex', flexDirection: 'column', gap: 16, color: '#f8fafc' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 'bold', color: '#38bdf8' }}>🎨 模板库与底图母版管理</div>
            <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 2 }}>已开启分辨率检测：推荐仅选用 ≥ 1920×1080 的高清底图，避免拉伸模糊</div>
          </div>
          <button onClick={onClose} style={{ background: '#334155', color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px', cursor: 'pointer' }}>关闭</button>
        </div>

        {/* 一、直接上传参考 .pptx 模板 */}
        <div style={{ padding: 14, background: '#0f172a', borderRadius: 8, border: '1px solid #38bdf8', display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 'bold', color: '#38bdf8' }}>一、直接上传参考 PPT 模板 (.pptx)</div>
              <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}>自动提取配色、过滤低清碎图，按清晰度从高到低展示可用底图</div>
            </div>
            <button onClick={() => pptxInputRef.current.click()} disabled={parsing} style={{ padding: '8px 14px', fontSize: 12, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold' }}>
              {parsing ? '解析中...' : '📤 上传 .pptx 模板'}
            </button>
            <input ref={pptxInputRef} type="file" accept=".pptx" onChange={handlePptxUpload} style={{ display: 'none' }} />
          </div>

          {pptxInfo && (
            <div style={{ padding: 10, background: '#1e293b', borderRadius: 6, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'center', fontSize: 12 }}>
                <span>已解析：<strong>{pptxInfo.fileName}</strong></span>
                {pptxInfo.primaryColor && (
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                    主色：<span style={{ width: 14, height: 14, borderRadius: 3, background: pptxInfo.primaryColor, display: 'inline-block', border: '1px solid #fff' }} />
                    {pptxInfo.primaryColor}
                  </span>
                )}
                {pptxInfo.accentColor && (
                  <button onClick={() => onApply({ customPrimary: pptxInfo.accentColor }, `已切换为模板辅色 ${pptxInfo.accentColor}`)} style={{ padding: '2px 8px', fontSize: 11, background: '#334155', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
                    改用辅色 ({pptxInfo.accentColor})
                  </button>
                )}
              </div>

              {pptxInfo.detectedFont && /微软雅黑|黑体|宋体|方正|汉仪|华康/.test(pptxInfo.detectedFont) && (
                <div style={{ fontSize: 11, color: '#fde047', background: '#422006', padding: '5px 8px', borderRadius: 4 }}>
                  🛡️ 合规拦截：原模板使用了「{pptxInfo.detectedFont}」，已自动锁定使用免费商用白名单字体！
                </div>
              )}

              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 11, color: '#94a3b8' }}>
                <span>可用高清/大图 ({visibleImages.length} 张){hiddenCount > 0 ? ` · 已自动隐藏 ${hiddenCount} 张低清小图标` : ''}</span>
                {hiddenCount > 0 && (
                  <label style={{ cursor: 'pointer', color: '#38bdf8' }}>
                    <input type="checkbox" checked={showSmallImgs} onChange={(e) => setShowSmallImgs(e.target.checked)} /> 显示小图
                  </label>
                )}
              </div>

              {visibleImages.length > 0 ? (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8, maxHeight: 220, overflowY: 'auto' }}>
                  {visibleImages.map((img, idx) => {
                    const q = getQualityTag(img.w, img.h);
                    return (
                      <div key={idx} style={{ background: '#0f172a', borderRadius: 6, padding: 5, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 4 }}>
                        <div style={{ fontSize: 10, padding: '2px 5px', borderRadius: 3, background: q.bg, color: q.color, textAlign: 'center', fontWeight: 'bold' }}>
                          {q.label}
                        </div>
                        <img src={img.dataUrl} alt={img.name} style={{ width: '100%', height: 68, objectFit: 'cover', borderRadius: 3 }} />
                        <div style={{ display: 'flex', gap: 4 }}>
                          <button onClick={() => onApply({ coverBg: img.dataUrl }, `已设为封面底图（${img.w}×${img.h}）`)} style={{ flex: 1, padding: '3px 0', fontSize: 10, background: '#1d4ed8', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer' }}>做封面</button>
                          <button onClick={() => onApply({ contentBg: img.dataUrl }, `已设为内页底图（${img.w}×${img.h}）`)} style={{ flex: 1, padding: '3px 0', fontSize: 10, background: '#0f766e', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer' }}>做内页</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div style={{ fontSize: 11, color: '#fbbf24', background: '#422006', padding: '7px 10px', borderRadius: 4 }}>
                  💡 该模板未包含高清背景位图（它的封面/内页是用 PPT 矢量形状拼出来的）。如需清晰复用其视觉，请在 WPS 中用「💡 该模板未包含高清背景位图（它的封面/内页是用 PPT 矢量形状拼出来的）。如需清晰复用其视觉，请在 WPS 中用「输出为图片 → 超清(300DPI)」→」导出无字底图，再在下方手动上传。
                </div>
              )}
            </div>
          )}
        </div>

        {/* 二、手动上传高清封面 / 内页底图 */}
        <div style={{ padding: 14, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>二、手动上传高清封面 / 内页底图（建议 1920×1080 PNG）</div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <div style={{ padding: 10, background: '#1e293b', borderRadius: 6, border: '1px dashed #475569', display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ fontSize: 12, color: '#cbd5e1', fontWeight: 'bold' }}>封面 / 章节 / 结尾底图</div>
              {config.coverBg ? (
                <img src={config.coverBg} alt="封面底图" style={{ width: '100%', height: 86, objectFit: 'cover', borderRadius: 4 }} />
              ) : (
                <div style={{ height: 86, background: '#0f172a', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, color: '#64748b' }}>使用默认高清矢量/渐变</div>
              )}
              <div style={{ display: 'flex', gap: 6 }}>
                <button onClick={() => coverInputRef.current.click()} style={{ flex: 1, padding: 6, fontSize: 11, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
                  {config.coverBg ? '更换底图' : '+ 上传高清封面图'}
                </button>
                {config.coverBg && <button onClick={() => onApply({ coverBg: '' }, '已恢复默认高清封面背景')} style={{ padding: '6px 8px', fontSize: 11, background: '#7f1d1d', color: '#fca5a5', border: 'none', borderRadius: 4, cursor: 'pointer' }}>清除</button>}
              </div>
              <input ref={coverInputRef} type="file" accept="image/*" onChange={handleBgUpload('coverBg')} style={{ display: 'none' }} />
            </div>

            <div style={{ padding: 10, background: '#1e293b', borderRadius: 6, border: '1px dashed #475569', display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ fontSize: 12, color: '#cbd5e1', fontWeight: 'bold' }}>内容页通用底图</div>
              {config.contentBg ? (
                <img src={config.contentBg} alt="内页底图" style={{ width: '100%', height: 86, objectFit: 'cover', borderRadius: 4 }} />
              ) : (
                <div style={{ height: 86, background: '#0f172a', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, color: '#64748b' }}>使用默认纯净白/渐变</div>
              )}
              <div style={{ display: 'flex', gap: 6 }}>
                <button onClick={() => contentInputRef.current.click()} style={{ flex: 1, padding: 6, fontSize: 11, background: '#0284c7', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
                  {config.contentBg ? '更换底图' : '+ 上传高清内页图'}
                </button>
                {config.contentBg && <button onClick={() => onApply({ contentBg: '' }, '已恢复默认内页背景')} style={{ padding: '6px 8px', fontSize: 11, background: '#7f1d1d', color: '#fca5a5', border: 'none', borderRadius: 4, cursor: 'pointer' }}>清除</button>}
              </div>
              <input ref={contentInputRef} type="file" accept="image/*" onChange={handleBgUpload('contentBg')} style={{ display: 'none' }} />
            </div>
          </div>
        </div>

        {/* 三、风格模板包导入 / 导出 */}
        <div style={{ padding: 14, background: '#0f172a', borderRadius: 8, border: '1px solid #334155', display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 'bold', color: '#e2e8f0' }}>三、积累自己的接单模板库（导出/导入 .json 模板包）</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <input value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="给模板起个名字，如：红蓝工作汇报模板" style={{ flex: 1, padding: '7px 10px', background: '#1e293b', color: '#fff', border: '1px solid #475569', borderRadius: 6, fontSize: 12 }} />
            <button onClick={handleExportTemplate} style={{ padding: '7px 14px', fontSize: 12, background: '#0f766e', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold' }}>💾 保存为模板包</button>
          </div>
          <button onClick={() => importTplRef.current.click()} style={{ padding: 10, fontSize: 13, background: '#2563eb', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold' }}>
            📂 导入已有模板包 (.json) 并一键换肤
          </button>
          <input ref={importTplRef} type="file" accept=".json" onChange={handleImportTemplate} style={{ display: 'none' }} />
        </div>
      </div>
    </div>
  );
}