// ==========================================
// Word 导出模块 (DOCX)
// ==========================================

// --- Word (.docx) 导出 (使用 docx 库) ---
async function exportToWord() {
    try {
        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
            WidthType, BorderStyle, AlignmentType, SectionType, convertInchesToTwip, LevelFormat } = docx;

        const contentNode = document.getElementById('book-content');
        if (!contentNode) throw new Error("Cannot find book content");

        const colors = getThemeColors();
        const accentHex = colors.accent.replace('#', '');
        const bgPanelHex = colors.bgPanel.replace('#', '');
        const textHex = colors.textPrimary.replace('#', '');
        const borderHex = colors.border.replace('#', '');
        const mutedHex = colors.textMuted.replace('#', '');

        // 获取字体 (从 CSS 变量提取)
        const fontStack = getCssVar('--font-serif');
        const fontSerif = fontStack.split(',')[0].replace(/["']/g, '').trim() || 'Georgia';
        // 根据字体栈类型确定中文字体 (匹配浏览器 fallback 行为)
        let fontChinese = 'SimSun'; // 默认宋体
        if (fontStack.includes('sans-serif')) {
            fontChinese = 'Microsoft YaHei'; // 微软雅黑
        } else if (fontStack.includes('monospace')) {
            fontChinese = 'SimSun'; // 等宽中文用宋体
        }

        // 提取内容
        const data = parseBookContent();

        // 辅助：获取 DOCX 字号 (1px = 1.5 half-points)
        // 动态读取 DOM 计算样式，实现 "所见即所得"
        // 修正：使用 Math.floor 向下取整，防止 Word 渲染比 CSS 宽
        function getDocxFontSize(selector, defaultSize) {
            const el = document.querySelector(selector);
            if (!el) return defaultSize;
            const fontSizePx = parseFloat(getComputedStyle(el).fontSize);
            if (isNaN(fontSizePx)) return defaultSize;
            return Math.floor(fontSizePx * 1.5);
        }

        // === 标题区内容 (单栏) ===
        const headerChildren = [];

        // 中文标题
        const titleSize = getDocxFontSize(SELECTORS.title, 60);
        headerChildren.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
            children: [new TextRun({ text: data.titleCn, bold: true, size: titleSize, color: textHex })]
        }));

        // 英文标题 + 下划线
        // 使用 mutedHex 配合 font-family:serif (subtitle 的 CSS 现在是 muted 色的)
        const subTitleSize = getDocxFontSize(SELECTORS.title + ' span', 28);
        headerChildren.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: textHex } },
            children: [new TextRun({ text: data.titleEn, size: subTitleSize, allCaps: true, color: mutedHex })]
        }));

        // 动态获取 Meta Gap
        const metaEl = document.querySelector('.book-meta');
        const metaGapPx = metaEl ? parseFloat(getComputedStyle(metaEl).gap) || 30 : 30;
        const halfGapTwips = Math.round(metaGapPx * 15 / 2); // 1px = 15 twips

        // 元信息表格 (时代 + Boss，AUTO 宽度居中)
        const metaTable = new Table({
            alignment: AlignmentType.CENTER,
            width: { size: 0, type: WidthType.AUTO },
            borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideHorizontal: { style: BorderStyle.NONE },
                insideVertical: { style: BorderStyle.NONE }
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: 0, type: WidthType.AUTO },
                            margins: { right: halfGapTwips },
                            children: [new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: `🕰️ ${data.era}`, size: 24, italics: true, color: mutedHex })]
                            })]
                        }),
                        new TableCell({
                            width: { size: 0, type: WidthType.AUTO },
                            margins: { left: halfGapTwips },
                            children: [new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: `💀 ${data.boss}`, size: 24, italics: true, color: mutedHex })]
                            })]
                        })
                    ]
                })
            ]
        });
        headerChildren.push(metaTable);

        // === 正文区内容 (双栏) ===
        const contentChildren = [];

        // 动态获取各类元素字号
        const h1Size = getDocxFontSize(SELECTORS.h1, 34);
        const pSize = getDocxFontSize(SELECTORS.p, 24);
        const sceneTitleSize = getDocxFontSize('.scene-title', 26);

        data.sections.forEach(section => {
            if (section.h1) {
                contentChildren.push(new Paragraph({
                    spacing: { before: 300, after: 100 },
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: borderHex } },
                    children: [new TextRun({ text: section.h1, bold: true, size: h1Size, color: accentHex })]
                }));
            }

            if (section.p) {
                contentChildren.push(new Paragraph({
                    spacing: { after: 200 },
                    children: [new TextRun({ text: section.p, size: pSize, color: textHex })]
                }));
            }

            // 时间轴
            section.timeline.forEach(item => {
                contentChildren.push(new Paragraph({
                    spacing: { after: 100 },
                    numbering: { reference: 'small-bullet', level: 0 },
                    children: [
                        new TextRun({ text: item.time, bold: true, color: accentHex, size: pSize }),
                        new TextRun({ text: item.text, size: pSize, color: textHex, break: 1 })
                    ]
                }));
            });

            // NPC 卡片 (根据用户规范)
            section.npcs.forEach(npc => {
                // 计算卡片背景色 (比 bgPanel 略深，模拟 rgba(0,0,0,0.04) 叠加效果)
                const bgR = parseInt(bgPanelHex.substring(0, 2), 16);
                const bgG = parseInt(bgPanelHex.substring(2, 4), 16);
                const bgB = parseInt(bgPanelHex.substring(4, 6), 16);
                const darkenFactor = 0.96;
                const cardBgHex = Math.round(bgR * darkenFactor).toString(16).padStart(2, '0') +
                    Math.round(bgG * darkenFactor).toString(16).padStart(2, '0') +
                    Math.round(bgB * darkenFactor).toString(16).padStart(2, '0');

                // 卡片内容
                const npcChildren = [];

                // 头像 (内嵌表格，与文字共享背景)
                const portraitTable = new Table({
                    width: { size: 1200, type: WidthType.DXA },
                    borders: {
                        top: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        bottom: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        left: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        right: { style: BorderStyle.SINGLE, size: 4, color: borderHex }
                    },
                    rows: [
                        new TableRow({
                            height: { value: 1400, rule: 'atLeast' },
                            children: [
                                new TableCell({
                                    shading: { fill: borderHex },
                                    verticalAlign: 'center',
                                    children: [new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: '?', size: 56, color: textHex })]
                                    })]
                                })
                            ]
                        })
                    ]
                });
                npcChildren.push(portraitTable);
                npcChildren.push(new Paragraph({ spacing: { after: 100 } }));

                // 名字 (强调色粗体 + 装饰线)
                npcChildren.push(new Paragraph({
                    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: borderHex } },
                    spacing: { after: 40 },
                    children: [new TextRun({ text: npc.name, bold: true, color: accentHex, size: pSize })]
                }));

                // 角色
                if (npc.role) {
                    npcChildren.push(new Paragraph({
                        children: [new TextRun({ text: npc.role, size: pSize, color: textHex })]
                    }));
                }

                // 属性
                if (npc.stats) {
                    npcChildren.push(new Paragraph({
                        children: [new TextRun({ text: npc.stats, size: pSize, color: textHex })]
                    }));
                }

                // 描述
                if (npc.desc) {
                    npcChildren.push(new Paragraph({
                        spacing: { before: 80 },
                        children: [new TextRun({ text: npc.desc, size: pSize, color: textHex })]
                    }));
                }

                // 秘密区域
                if (npc.secret) {
                    npcChildren.push(new Paragraph({
                        spacing: { before: 80 },
                        children: [
                            new TextRun({ text: '⚠ 秘密：', bold: true, color: accentHex, size: pSize }),
                            new TextRun({ text: npc.secret, size: pSize, color: textHex })
                        ]
                    }));
                }

                // NPC 卡片外框
                const npcTable = new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        bottom: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        left: { style: BorderStyle.SINGLE, size: 4, color: borderHex },
                        right: { style: BorderStyle.SINGLE, size: 4, color: borderHex }
                    },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    shading: { fill: cardBgHex },
                                    margins: { top: 150, bottom: 150, left: 200, right: 200 },
                                    children: npcChildren
                                })
                            ]
                        })
                    ]
                });
                contentChildren.push(npcTable);
                contentChildren.push(new Paragraph({ spacing: { after: 300 } }));
            });

            // 场景卡片
            section.scenes.forEach(scene => {
                // 场景背景色 (bgPanel * 0.98，模拟 rgba(0,0,0,0.02))
                const sceneDarken = 0.98;
                const sceneBgHex = Math.round(parseInt(bgPanelHex.substring(0, 2), 16) * sceneDarken).toString(16).padStart(2, '0') +
                    Math.round(parseInt(bgPanelHex.substring(2, 4), 16) * sceneDarken).toString(16).padStart(2, '0') +
                    Math.round(parseInt(bgPanelHex.substring(4, 6), 16) * sceneDarken).toString(16).padStart(2, '0');

                const sceneTable = new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.NONE },
                        bottom: { style: BorderStyle.NONE },
                        left: { style: BorderStyle.THICK, size: 24, color: mutedHex },
                        right: { style: BorderStyle.NONE }
                    },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    shading: { fill: sceneBgHex },
                                    margins: { top: 100, bottom: 100, left: 150, right: 150 },
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({ text: scene.title, bold: true, size: sceneTitleSize, color: textHex }),
                                                scene.item ? new TextRun({ text: '  ' + scene.item, size: pSize, color: mutedHex }) : null
                                            ].filter(Boolean)
                                        }),
                                        new Paragraph({ spacing: { before: 100 }, children: [new TextRun({ text: scene.desc, size: pSize, color: textHex })] }),
                                        scene.event ? new Paragraph({ spacing: { before: 100 }, children: [new TextRun({ text: scene.event, italics: true, bold: true, size: pSize, color: accentHex })] }) : new Paragraph({})
                                    ]
                                })
                            ]
                        })
                    ]
                });
                contentChildren.push(sceneTable);
                contentChildren.push(new Paragraph({ spacing: { after: 200 } }));
            });
        });

        // 页面设置 (两个 section 共享)
        const pageSettings = {
            // 1:1 Pixel Mapping (96 DPI): 800px Width, 60px Padding (Back to 1:1 based on Math.floor fix)
            size: { width: 12000, height: convertInchesToTwip(11.69) }, // 800px
            margin: { top: convertInchesToTwip(1), bottom: convertInchesToTwip(0.5), left: 900, right: 900 } // 60px (900 twips)
        };

        // 创建文档 (两个 section，共享页面设置以确保 CONTINUOUS 正常工作)
        const doc = new Document({
            background: { color: bgPanelHex },
            styles: {
                default: {
                    document: {
                        run: {
                            characterSpacing: -5, // Condense by 0.25pt to match CSS tightness
                            font: {
                                ascii: fontSerif,
                                eastAsia: fontChinese,
                                hAnsi: fontSerif
                            },
                            size: 24 // Default size 24 (12pt = 16px)
                        }
                    }
                }
            },
            numbering: {
                config: [{
                    reference: 'small-bullet',
                    levels: [{
                        level: 0,
                        format: LevelFormat.BULLET,
                        text: '•',
                        alignment: AlignmentType.LEFT,
                        style: {
                            run: { size: pSize }, // Match paragraph font size (CSS default)
                            paragraph: {
                                indent: { left: 360, hanging: 180 } // Tighter bullet-text spacing
                            }
                        }
                    }]
                }]
            },
            sections: [
                {
                    // 标题区：单栏
                    properties: {
                        page: pageSettings,
                        column: { count: 1 }
                    },
                    children: headerChildren
                },
                {
                    // 正文区：双栏，紧接标题 (同一页)
                    properties: {
                        type: SectionType.CONTINUOUS,
                        page: pageSettings,
                        column: { count: 2, space: 600 } // 40px Gap
                    },
                    children: contentChildren
                }
            ]
        });

        const blob = await Packer.toBlob(doc);
        const fileTitle = document.getElementById('val-final-branch')?.getAttribute('data-title-cn') || 'ArkhamModule';
        downloadFile(blob, `${fileTitle}.docx`);

    } catch (e) {
        console.error(e);
        alert('Word Export Failed: ' + e.message);
    }
}
