// ==========================================
// 导出引擎工具集 Export Utils
// ==========================================

// 辅助：获取 CSS 变量值
function getCssVar(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}

// RGB/RGBA 转 Hex
function rgbToHex(color) {
    if (!color || color === 'transparent' || color === 'rgba(0, 0, 0, 0)') return '#ffffff';
    if (color.startsWith('#')) return color;
    const match = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (match) {
        return '#' + [match[1], match[2], match[3]]
            .map(x => parseInt(x).toString(16).padStart(2, '0')).join('');
    }
    return '#000000';
}

// 获取当前主题颜色
function getThemeColors() {
    const computedBody = getComputedStyle(document.body);
    return {
        accent: rgbToHex(computedBody.getPropertyValue('--accent').trim()) || '#8a0303',
        bgPanel: rgbToHex(computedBody.getPropertyValue('--bg-panel').trim()) || '#fdf6e3',
        textPrimary: rgbToHex(computedBody.getPropertyValue('--text-primary').trim()) || '#433422',
        textMuted: rgbToHex(computedBody.getPropertyValue('--text-muted').trim()) || '#8b7d6b',
        border: rgbToHex(computedBody.getPropertyValue('--border').trim()) || '#d4c5a9'
    };
}

function escapeTex(str) {
    if (!str) return '';
    return str.replace(/([&%$#_{}])/g, '\\$1').replace(/\n/g, ' ').trim();
}

// 解析书籍内容
function parseBookContent() {
    const contentNode = document.getElementById('book-content');
    if (!contentNode) throw new Error("Cannot find book content");

    const titleCn = contentNode.querySelector(SELECTORS.title)?.childNodes[0]?.textContent?.trim() || '未命名模组';
    const titleEn = contentNode.querySelector(SELECTORS.title + ' span')?.textContent?.trim() || 'Untitled Module';

    const metaSpans = contentNode.querySelectorAll(SELECTORS.meta + ' span');
    const era = metaSpans[0]?.textContent?.replace('🕰️', '').trim() || '';
    const boss = metaSpans[1]?.textContent?.replace('💀', '').trim() || '';

    const sections = [];
    contentNode.querySelectorAll(SELECTORS.section + ':not(' + SELECTORS.headerSection + ')').forEach(section => {
        const secData = {
            h1: section.querySelector(SELECTORS.h1)?.textContent || '',
            p: section.querySelector(SELECTORS.p)?.textContent || '',
            timeline: [],
            npcs: [],
            scenes: []
        };

        // Timeline
        const timelineUl = section.querySelector(SELECTORS.timeline.item)?.parentElement; // Usually UL inside section
        // Or if we strictly follow structure:
        const tList = section.querySelector('ul');
        if (tList) {
            tList.querySelectorAll(SELECTORS.timeline.item).forEach(li => {
                const strong = li.querySelector('strong');
                const timeText = strong ? strong.textContent : '';
                const restText = li.textContent.replace(timeText, '').trim();
                secData.timeline.push({ time: timeText, text: restText });
            });
        }

        // NPCs
        section.querySelectorAll(SELECTORS.npc.container).forEach(card => {
            const fullText = card.querySelector(SELECTORS.npc.desc)?.textContent || '';
            let desc = fullText, secret = '';
            const secretIdx = fullText.indexOf('秘密：');
            if (secretIdx > -1) {
                desc = fullText.substring(0, secretIdx).replace('⚠️', '').trim();
                secret = fullText.substring(secretIdx + 3).trim();
            }

            secData.npcs.push({
                name: card.querySelector('.npc-name')?.textContent || 'NPC', // Keep legacy for simplicity or add to SELECTORS if critical
                role: card.querySelector('.npc-role')?.textContent || '',
                stats: card.querySelector('.npc-stats')?.textContent || '',
                desc: desc,
                secret: secret
            });
        });

        // Scenes
        section.querySelectorAll(SELECTORS.scene.container).forEach(box => {
            secData.scenes.push({
                title: box.querySelector('.scene-title')?.textContent || '',
                item: box.querySelector('.scene-item')?.textContent || '',
                desc: box.querySelector(SELECTORS.scene.desc)?.textContent || '',
                event: box.querySelector(SELECTORS.scene.event)?.textContent || ''
            });
        });

        sections.push(secData);
    });

    return { titleCn, titleEn, era, boss, sections };
}
