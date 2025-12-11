// ==========================================
// LaTeX 导出模块 (.tex)
// ==========================================

// --- LaTeX (.tex) 导出 (XeLaTeX + CJK) ---
async function exportToTex() {
    try {
        const data = parseBookContent();

        const colors = getThemeColors();
        const pureAccent = colors.accent.replace('#', '');
        const pureBgPanel = colors.bgPanel.replace('#', '');
        const pureBorder = colors.border.replace('#', '');
        const pureText = colors.textPrimary.replace('#', '');
        const pureMuted = colors.textMuted.replace('#', '');
        // 获取字体 (从 CSS 变量提取，与 DOCX 导出一致)
        const fontStack = getCssVar('--font-serif');
        const fontSerif = fontStack.split(',')[0].replace(/["']/g, '').trim() || 'Times New Roman';
        // 根据字体栈类型确定中文字体
        let fontChinese = 'FandolSong'; // 默认宋体 (Overleaf 内置)
        if (fontStack.includes('sans-serif')) {
            fontChinese = 'Noto Sans CJK SC';
        }

        // 辅助：处理文本中的 Emoji -> FontAwesome
        function processTexText(str) {
            let processed = escapeTex(str);
            // 常见 Emoji 替换列表 (XeLaTeX 默认不支持彩色 Emoji，用 FontAwesome 矢量图标代替)
            const map = {
                '🕰️': '\\faClock',
                '🕰': '\\faClock',
                '💀': '\\faSkull',
                '⚠️': '\\faExclamationTriangle',
                '⚡': '\\faBolt',
                '📍': '\\faMapMarker*',
                '📦': '\\faBox',
                '🕯️': '\\faBurn',
                '🕯': '\\faBurn',
                '🔎': '\\faSearch',
                '🔍': '\\faSearch',
                '🖋️': '\\faPenFancy',
                '🖋': '\\faPenFancy',
                '👥': '\\faUsers',
                '📄': '\\faFileAlt',
                '💾': '\\faSave',
                '📐': '\\faRulerCombined',
                '🖨️': '\\faPrint',
                '🖨': '\\faPrint'
            };

            for (const [emoji, icon] of Object.entries(map)) {
                // 全局替换
                processed = processed.split(emoji).join(` {\\small ${icon}} `);
            }
            return processed;
        }

        let tex = `% 使用 XeLaTeX 编译 (Overleaf: Menu -> Compiler -> XeLaTeX)
\\documentclass[a4paper, 10pt]{article}

% CJK 支持 (XeLaTeX)
\\usepackage{fontspec}
\\setmainfont{${fontSerif}}
\\usepackage{xeCJK}
\\setCJKmainfont[AutoFakeSlant=0.2]{${fontChinese}}

% 图标支持 (Emoji 替代品)
\\usepackage{fontawesome5}

% 页面设置
\\usepackage[margin=2cm]{geometry}
\\usepackage{multicol}
\\setlength{\\columnsep}{1cm}

% 颜色
\\usepackage{xcolor}
\\usepackage{pagecolor}
\\definecolor{accent}{HTML}{${pureAccent}}
\\definecolor{pagebg}{HTML}{${pureBgPanel}}
\\definecolor{bordercolor}{HTML}{${pureBorder}}
\\definecolor{textcolor}{HTML}{${pureText}}
\\definecolor{mutedcolor}{HTML}{${pureMuted}}
\\pagecolor{pagebg}
\\color{textcolor}

% 卡片样式
\\usepackage[most]{tcolorbox}

% NPC 卡片 - 紧凑垂直布局
\\newtcolorbox{npccard}{
  enhanced,
  breakable,
  colback=pagebg!96!black,
  colframe=bordercolor,
  boxrule=0.5pt,
  arc=0pt,
  left=4pt, right=4pt, top=4pt, bottom=4pt,
  before skip=6pt,
  after skip=6pt
}

% 场景卡片
\\newtcolorbox{scenebox}{
  enhanced,
  breakable,
  frame hidden,
  colback=pagebg!98!black,
  borderline west={3pt}{0pt}{mutedcolor},
  left=8pt, right=6pt, top=4pt, bottom=4pt,
  before skip=6pt,
  after skip=6pt
}

% 章节标题 (使用 parbox 确保标题和装饰线不分离，即使在多栏模式下)
\\newcommand{\\mysection}[1]{%
  \\vspace{1em}%
  \\noindent\\parbox{\\linewidth}{%
    {\\Large\\bfseries\\color{accent}#1}\\\\[-0.2em]%
    {\\color{bordercolor}\\rule{\\linewidth}{0.5pt}}%
  }%
  \\vspace{0.8em}\\par%
}

% 列表
\\usepackage{enumitem}
\\setlist[itemize]{leftmargin=*, itemsep=0.3em}

% 无页码
\\pagestyle{empty}

\\begin{document}

% === 标题区 ===
\\begin{center}
{\\Huge\\bfseries ${processTexText(data.titleCn)}}

\\vspace{0.3em}
{\\large\\scshape\\color{mutedcolor} ${processTexText(data.titleEn)}}
\\end{center}

\\vspace{0.3em}
{\\color{bordercolor}\\hrule height 0.5pt}
\\vspace{0.5em}

\\begin{center}
{\\color{mutedcolor}\\faClock\\ ${processTexText(data.era)} \\hspace{2em} \\faSkull\\ ${processTexText(data.boss)}}
\\end{center}

\\vspace{1em}

% === 正文区 (双栏) ===
\\begin{multicols}{2}
`;

        data.sections.forEach(section => {
            if (section.h1) tex += `\\mysection{${processTexText(section.h1)}}\n\n`;

            if (section.p) tex += `${processTexText(section.p)}\n\n`;

            // 时间轴
            if (section.timeline.length > 0) {
                tex += `\\begin{itemize}\n`;
                section.timeline.forEach(item => {
                    tex += `\\item {\\bfseries\\color{accent}${processTexText(item.time)}} ${processTexText(item.text)}\n`;
                });
                tex += `\\end{itemize}\n\n`;
            }

            // NPC 卡片 - 垂直布局
            section.npcs.forEach(npc => {
                tex += `\\begin{npccard}\n`;
                tex += `\\fcolorbox{textcolor}{bordercolor}{\\parbox[c][1.6cm][c]{1.3cm}{\\centering\\fontsize{24pt}{28pt}\\selectfont\\rmfamily\\color{textcolor} ?}}\n\n`;
                tex += `\\vspace{0.2em}\n`;
                tex += `{\\bfseries\\color{accent}${processTexText(npc.name)}}\n\n`;
                tex += `{\\color{bordercolor}\\rule{\\linewidth}{0.3pt}}\n\n`;
                if (npc.role) tex += `{\\slshape\\small\\color{mutedcolor}${processTexText(npc.role)}}\n\n`;
                if (npc.stats) tex += `{\\footnotesize ${processTexText(npc.stats)}}\n\n`;
                if (npc.desc) tex += `{\\small ${processTexText(npc.desc)}}\n\n`;
                if (npc.secret) {
                    tex += `{\\small\\bfseries\\color{accent}\\faExclamationTriangle\\ 秘密：}{\\small ${processTexText(npc.secret)}}\n`;
                }
                tex += `\\end{npccard}\n\n`;
            });

            // 场景卡片
            section.scenes.forEach(scene => {
                tex += `\\begin{scenebox}\n`;
                tex += `{\\bfseries ${processTexText(scene.title)}}`;
                if (scene.item) tex += ` \\hfill {\\small\\color{mutedcolor}${processTexText(scene.item)}}`;
                tex += `\\par\\vspace{0.2em}\n`;
                tex += `${processTexText(scene.desc)}\\par\\vspace{0.2em}\n`;
                if (scene.event) {
                    tex += `{\\slshape\\bfseries\\color{accent}${processTexText(scene.event)}}\n`;
                }
                tex += `\\end{scenebox}\n\n`;
            });
        });

        tex += `\\end{multicols}\n\\end{document}`;

        const blob = new Blob([tex], { type: 'text/plain;charset=utf-8' });
        const fileTitle = document.getElementById('val-final-branch')?.getAttribute('data-title-cn') || 'ArkhamModule';
        downloadFile(blob, `${fileTitle}.tex`);

    } catch (e) {
        console.error(e);
        alert('TeX Export Failed: ' + e.message);
    }
}
