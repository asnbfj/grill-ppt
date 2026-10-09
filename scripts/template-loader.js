#!/usr/bin/env node

/**
 * template-loader - 用户自定义 PPT 母版加载器
 *
 * 职责：
 *   1. 扫描 templates/ 目录，结合 registry.json 列出可用模板
 *   2. 解析 .pptx（zip）里的 ppt/theme/theme1.xml，提取主题配色与字体
 *   3. 把主题映射成 generate.js 可直接使用的 palette 结构
 *
 * 设计约束：
 *   - 零新增 npm 依赖：解压复用系统 unzip（macOS / Linux 自带）
 *   - pptxgenjs 无法以已有 .pptx 为基底追加页面，故采用「主题提取」方案，
 *     详见 templates/README.md 的「实现边界」
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const TEMPLATES_DIR = path.join(PROJECT_ROOT, 'templates');
const REGISTRY_FILE = path.join(TEMPLATES_DIR, 'registry.json');

const DEFAULT_FONT = 'Microsoft YaHei';

/* ------------------------------------------------------------------ *
 * 目录与登记表
 * ------------------------------------------------------------------ */

function ensureTemplatesDir() {
  if (!fs.existsSync(TEMPLATES_DIR)) {
    fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
  }
  return TEMPLATES_DIR;
}

function readRegistry() {
  try {
    const parsed = JSON.parse(fs.readFileSync(REGISTRY_FILE, 'utf8'));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    // registry.json 缺失或非法 JSON 都不应阻塞生成，退化为「无元数据」
    return {};
  }
}

/**
 * 列出 templates/ 下所有可用模板
 * @returns {Array<{file:string, path:string, name:string, style:string, scene:string, desc:string}>}
 */
function listTemplates() {
  ensureTemplatesDir();
  const registry = readRegistry();

  return fs
    .readdirSync(TEMPLATES_DIR)
    .filter((f) => {
      const lower = f.toLowerCase();
      return lower.endsWith('.pptx') && !f.startsWith('~$');
    })
    .sort()
    .map((file) => {
      const meta = registry[file] || {};
      return {
        file,
        path: path.join(TEMPLATES_DIR, file),
        name: meta.name || path.basename(file, path.extname(file)),
        style: meta.style || '',
        scene: meta.scene || '',
        desc: meta.desc || ''
      };
    });
}

/**
 * 把「模板引用」解析成磁盘上的绝对路径。
 * 支持三种写法：文件名（含/不含扩展名）、登记的中文名、任意路径（含 ~）。
 */
function resolveTemplate(ref) {
  if (!ref || typeof ref !== 'string') {
    throw new Error('未提供模板名称或路径（--template 的值为空）');
  }

  const expanded = ref.startsWith('~')
    ? path.join(process.env.HOME || '', ref.slice(1))
    : ref;

  const candidates = [
    expanded,
    path.resolve(process.cwd(), expanded),
    path.join(TEMPLATES_DIR, expanded)
  ];

  for (const candidate of candidates) {
    try {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()
          && candidate.toLowerCase().endsWith('.pptx')) {
        return candidate;
      }
    } catch (err) {
      /* 忽略无权限等个别路径，继续尝试下一个 */
    }
  }

  const templates = listTemplates();
  const hit = templates.find((t) =>
    t.file === ref ||
    t.name === ref ||
    path.basename(t.file, path.extname(t.file)) === ref
  );
  if (hit) return hit.path;

  const available = templates.length
    ? templates.map((t) => `${t.name} (${t.file})`).join('、')
    : '（templates/ 目录为空，请先放入 .pptx 模板）';
  throw new Error(`找不到模板「${ref}」。可用模板：${available}`);
}

/* ------------------------------------------------------------------ *
 * zip 读取（复用系统 unzip）
 * ------------------------------------------------------------------ */

function zipList(file) {
  return execFileSync('unzip', ['-Z1', file], { encoding: 'utf8', maxBuffer: 1 << 26 })
    .split('\n')
    .map((s) => s.trim())
    .filter(Boolean);
}

function zipRead(file, entry) {
  return execFileSync('unzip', ['-p', file, entry], { encoding: 'utf8', maxBuffer: 1 << 26 });
}

/* ------------------------------------------------------------------ *
 * theme1.xml 解析
 * ------------------------------------------------------------------ */

/**
 * 从 clrScheme 片段里取某个颜色槽（dk1/lt1/accent1...）。
 * Office 主题里颜色可能是 srgbClr、也可能是 sysClr + lastClr，两种都要认。
 */
function extractColor(clrSchemeXml, key) {
  const block = new RegExp(`<a:${key}\\b[^>]*>([\\s\\S]*?)</a:${key}>`).exec(clrSchemeXml);
  if (!block) return null;
  const inner = block[1];

  const srgb = /<a:srgbClr[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner);
  if (srgb) return srgb[1].toUpperCase();

  const lastClr = /\blastClr="([0-9A-Fa-f]{6})"/.exec(inner);
  if (lastClr) return lastClr[1].toUpperCase();

  const sysVal = /<a:sysClr[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner);
  if (sysVal) return sysVal[1].toUpperCase();

  return null;
}

const COLOR_SLOTS = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink'
];

function parseTheme(themeXml) {
  const clrScheme = (/<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(themeXml) || [])[0] || '';
  const fontScheme = (/<a:fontScheme[\s\S]*?<\/a:fontScheme>/.exec(themeXml) || [])[0] || '';

  const colors = {};
  if (clrScheme) {
    for (const slot of COLOR_SLOTS) {
      const value = extractColor(clrScheme, slot);
      if (value) colors[slot] = value;
    }
  }

  const fonts = {};
  if (fontScheme) {
    const major = (/<a:majorFont>[\s\S]*?<\/a:majorFont>/.exec(fontScheme) || [])[0] || '';
    const minor = (/<a:minorFont>[\s\S]*?<\/a:minorFont>/.exec(fontScheme) || [])[0] || '';
    // 拉丁字体决定生成时的 fontFace；东亚字体无法映射到 pptxgenjs 的单一 fontFace
    const latinOf = (block) => (/<a:latin[^>]*\btypeface="([^"]*)"/.exec(block) || [])[1] || null;
    fonts.major = latinOf(major);
    fonts.minor = latinOf(minor);
  }

  const name = (/<a:theme[^>]*\bname="([^"]*)"/.exec(themeXml) || [])[1] || '';
  return { name, colors, fonts };
}

/* ------------------------------------------------------------------ *
 * 对比度保护（呼应 SKILL.md「不让文字和背景对比度不足」）
 * ------------------------------------------------------------------ */

function relativeLuminance(hex) {
  const channels = [0, 2, 4]
    .map((i) => parseInt(hex.substr(i, 2), 16) / 255)
    .map((v) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)));
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

function contrastRatio(a, b) {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  const hi = Math.max(la, lb);
  const lo = Math.min(la, lb);
  return (hi + 0.05) / (lo + 0.05);
}

/**
 * 保证 color 在 bg 上可读；不达标时按顺序尝试 fallbacks，最后按底色明暗兜底。
 */
function ensureContrast(color, bg, fallbacks = [], min = 3) {
  const usable = (c) => c && /^[0-9A-Fa-f]{6}$/.test(c);
  if (usable(color) && usable(bg) && contrastRatio(color, bg) >= min) return color.toUpperCase();
  for (const fb of fallbacks) {
    if (usable(fb) && usable(bg) && contrastRatio(fb, bg) >= min) return fb.toUpperCase();
  }
  return relativeLuminance(bg) > 0.5 ? '1F2937' : 'F0F6FC';
}

/* ------------------------------------------------------------------ *
 * 主题 → palette
 * ------------------------------------------------------------------ */

function buildPalette(theme, name) {
  const c = theme.colors || {};
  const pick = (...keys) => {
    for (const k of keys) if (c[k]) return c[k];
    return null;
  };

  const bgLight = pick('lt1') || 'FFFFFF';
  const bgDark = pick('dk2', 'dk1') || '0D1117';

  const primary = ensureContrast(
    pick('accent1', 'dk2', 'dk1') || '1E2761',
    bgLight,
    [pick('accent1'), pick('accent2'), pick('accent3'), pick('dk2'), pick('dk1')],
    3
  );
  const text = ensureContrast(
    pick('dk1', 'dk2') || '333333',
    bgLight,
    [pick('dk2'), pick('dk1'), pick('accent1')],
    4.5
  );

  const palette = {
    name,
    source: 'template',
    primary,
    // secondary / accent 都渲染在 primary 底色上（封面副标题、数据卡标签、总结序号圆点），需保证可读
    secondary: ensureContrast(
      pick('accent2', 'lt2', 'accent1') || 'CADCFC',
      primary,
      [pick('accent2'), pick('lt2'), pick('lt1'), pick('accent3'), pick('accent4')],
      3
    ),
    accent: ensureContrast(
      pick('accent3', 'accent4', 'accent1', 'lt1') || 'FFFFFF',
      primary,
      [pick('accent3'), pick('accent4'), pick('lt1'), pick('accent2')],
      3
    ),
    bgLight,
    bgDark,
    text,
    textLight: ensureContrast(pick('dk2', 'dk1') || '6B7280', bgLight, ['6B7280'], 3),
    onPrimary: ensureContrast('FFFFFF', primary, [pick('accent2'), pick('accent3'), pick('lt1')], 4.5),
    fontFace: (theme.fonts && (theme.fonts.major || theme.fonts.minor)) || DEFAULT_FONT
  };

  return palette;
}

/* ------------------------------------------------------------------ *
 * 对外主入口
 * ------------------------------------------------------------------ */

/**
 * 加载模板：解析引用 → 读主题 → 产出 palette
 * @param {string} ref 模板名 / 文件名 / 路径
 * @returns {{file:string, name:string, themeEntry:string, colors:object, fonts:object, palette:object, meta:object}}
 */
function loadTemplate(ref) {
  const file = resolveTemplate(ref);
  const meta = listTemplates().find((t) => t.path === file) || {};
  const displayName = meta.name || path.basename(file, path.extname(file));

  let entries;
  try {
    entries = zipList(file);
  } catch (err) {
    throw new Error(
      `无法读取模板文件「${file}」。请确认它是有效的 .pptx，且系统已安装 unzip。\n原始错误：${err.message}`
    );
  }

  const themeEntry = entries.find((e) => /^ppt\/theme\/theme\d+\.xml$/i.test(e));
  if (!themeEntry) {
    throw new Error(`模板「${file}」中未找到主题文件（ppt/theme/theme*.xml），可能不是标准 .pptx。`);
  }

  const theme = parseTheme(zipRead(file, themeEntry));
  const palette = buildPalette(theme, displayName);

  return {
    file,
    name: displayName,
    themeEntry,
    themeName: theme.name,
    colors: theme.colors,
    fonts: theme.fonts,
    palette,
    meta
  };
}

module.exports = {
  TEMPLATES_DIR,
  REGISTRY_FILE,
  DEFAULT_FONT,
  listTemplates,
  resolveTemplate,
  parseTheme,
  buildPalette,
  loadTemplate
};
