// ── NG Rate Summary Chart ─────────────────────────────────────────────────────
window.ngRateChart = {
    _instances: {},

    render: function (canvasId, labels, barDatasets, totalDataset) {
        if (this._instances[canvasId]) {
            this._instances[canvasId].destroy();
            delete this._instances[canvasId];
        }
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const datasets = [...barDatasets];
        if (totalDataset) {
            datasets.push(Object.assign({ type: 'line' }, totalDataset));
        }

        // Set of separator indices (positions where the label is "")
        const sepIdx = new Set(labels.reduce((acc, l, i) => { if (l === '') acc.push(i); return acc; }, []));

        this._instances[canvasId] = new Chart(canvas, {
            type: 'bar',
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 12, padding: 10, font: { size: 10 } }
                    },
                    tooltip: {
                        filter: item => !sepIdx.has(item.dataIndex),
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y != null ? ctx.parsed.y.toLocaleString() : '-'}`
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        ticks: {
                            font: { size: 9 },
                            maxRotation: 0,
                            callback: function(val, i) { return sepIdx.has(i) ? '' : labels[i]; }
                        },
                        grid: { color: ctx => sepIdx.has(ctx.index) ? 'transparent' : '#f1f5f9' }
                    },
                    y: {
                        stacked: true,
                        ticks: { font: { size: 9 }, callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v },
                        grid: { color: '#e2e8f0' }
                    },
                    y1: {
                        position: 'right',
                        ticks: { font: { size: 9 }, callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v },
                        grid: { drawOnChartArea: false }
                    }
                }
            }
        });
    },

    destroy: function (canvasId) {
        if (this._instances[canvasId]) {
            this._instances[canvasId].destroy();
            delete this._instances[canvasId];
        }
    }
};

// ── Auto-reload on Blazor reconnect ──────────────────────────────────────────
// Blazor Server는 서버 재시작 시 "Attempting to reconnect" 모달을 띄우고 기본 8회 재시도 후 실패하면 멈춥니다.
// 우리는 재연결 모달이 뜨는 즉시 주기적으로 서버를 핑하고, 응답이 오면 페이지를 하드 리로드해서
// 사용자가 직접 새로고침하지 않아도 새 빌드가 자동 반영되게 합니다.
document.addEventListener('DOMContentLoaded', () => {
    const modal = document.getElementById('components-reconnect-modal');
    if (!modal) return;

    let pollTimer = null;
    const startPoll = () => {
        if (pollTimer) return;
        pollTimer = setInterval(async () => {
            try {
                const res = await fetch(window.location.pathname, {
                    method: 'HEAD',
                    cache:  'no-store',
                });
                if (res.ok) {
                    clearInterval(pollTimer);
                    location.reload();
                }
            } catch (_) { /* server still down */ }
        }, 1500);
    };
    const stopPoll = () => {
        if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
    };

    const shouldPoll = () =>
        modal.classList.contains('components-reconnect-show')     ||
        modal.classList.contains('components-reconnect-failed')   ||
        modal.classList.contains('components-reconnect-rejected');

    new MutationObserver(() => {
        if (shouldPoll()) startPoll(); else stopPoll();
    }).observe(modal, { attributes: true, attributeFilter: ['class', 'style'] });
});

// ── NG Rate By-Group Line Chart ───────────────────────────────────────────────
window.ngRateGroupChart = {
    _instances: {},
    _palette: [
        '#6366f1', '#14b8a6', '#f97316', '#ef4444', '#8b5cf6',
        '#0ea5e9', '#f59e0b', '#10b981', '#ec4899', '#64748b',
        '#a855f7', '#22c55e', '#eab308', '#06b6d4', '#dc2626'
    ],

    render: function (canvasId, labels, series) {
        if (this._instances[canvasId]) {
            this._instances[canvasId].destroy();
            delete this._instances[canvasId];
        }
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        // Separator indices: positions where label === '' (used to visually divide Date / Week / Month blocks).
        const sepIdx = new Set(labels.reduce((acc, l, i) => { if (l === '') acc.push(i); return acc; }, []));

        const datasets = series.map((s, i) => {
            const color = this._palette[i % this._palette.length];
            return {
                label: s.name,
                data: s.values,
                borderColor: color,
                backgroundColor: color + '22',
                borderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 5,
                spanGaps: false, // don't bridge across separators / missing points
                tension: 0.25,
            };
        });

        this._instances[canvasId] = new Chart(canvas, {
            type: 'line',
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 12, padding: 8, font: { size: 10 } }
                    },
                    tooltip: {
                        filter: item => !sepIdx.has(item.dataIndex),
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y != null ? ctx.parsed.y.toLocaleString() : '-'}`
                        }
                    }
                },
                scales: {
                    x: {
                        ticks: {
                            font: { size: 10 },
                            maxRotation: 0,
                            callback: function (val, i) { return sepIdx.has(i) ? '' : labels[i]; }
                        },
                        grid: { color: ctx => sepIdx.has(ctx.index) ? '#94a3b8' : '#f1f5f9' }
                    },
                    y: {
                        ticks: { font: { size: 10 }, callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v },
                        grid: { color: '#e2e8f0' },
                        beginAtZero: true,
                    }
                }
            }
        });
    },

    destroy: function (canvasId) {
        if (this._instances[canvasId]) {
            this._instances[canvasId].destroy();
            delete this._instances[canvasId];
        }
    }
};

// ── File Download ─────────────────────────────────────────────────────────────
window.downloadBase64File = function (filename, base64, contentType) {
    const link = document.createElement('a');
    link.href     = `data:${contentType};base64,${base64}`;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

// ── Data Inference Drag-and-Drop ─────────────────────────────────────────────
window.diDragDrop = {
    _dotnet: null,
    _kind:   null,
    _idx:    -1,

    init: function (dotnetRef) {
        this._dotnet = dotnetRef;
        if (this._bound) return;
        this._bound = true;

        var self = this;

        document.addEventListener('dragstart', function (e) {
            var el = e.target.closest('[data-di-kind]');
            if (!el) return;
            self._kind = el.dataset.diKind;
            self._idx  = parseInt(el.dataset.diIdx) || 0;
            e.dataTransfer.effectAllowed = 'copy';
            e.dataTransfer.setData('text/plain', self._kind + ':' + self._idx);
        });

        document.addEventListener('dragover', function (e) {
            var zone = e.target.closest('[data-di-zone]');
            if (!zone) return;
            e.preventDefault();
            e.dataTransfer.dropEffect = 'copy';
            document.querySelectorAll('.di-dz-active')
                .forEach(function (z) { if (z !== zone) z.classList.remove('di-dz-active'); });
            zone.classList.add('di-dz-active');
        });

        document.addEventListener('dragleave', function (e) {
            var zone = e.target.closest('[data-di-zone]');
            if (!zone) return;
            if (e.relatedTarget && zone.contains(e.relatedTarget)) return;
            zone.classList.remove('di-dz-active');
        });

        document.addEventListener('drop', function (e) {
            var zone = e.target.closest('[data-di-zone]');
            document.querySelectorAll('.di-dz-active')
                .forEach(function (z) { z.classList.remove('di-dz-active'); });
            if (!zone) return;
            e.preventDefault();
            if (self._kind && self._dotnet) {
                var zoneIdx = parseInt(zone.dataset.diZone);
                self._dotnet.invokeMethodAsync('JsDrop', self._kind, self._idx, zoneIdx);
            }
            self._kind = null; self._idx = -1;
        });

        document.addEventListener('dragend', function () {
            document.querySelectorAll('.di-dz-active')
                .forEach(function (z) { z.classList.remove('di-dz-active'); });
            self._kind = null; self._idx = -1;
        });
    },

    destroy: function (ref) {
        // Only null out if the ref being destroyed is still the active one
        if (!ref || this._dotnet === ref) this._dotnet = null;
    }
};

// ── Data Inference Block Resize ──────────────────────────────────────────────
window.diResize = {
    _dotnet: null,

    init: function (dotnetRef) {
        this._dotnet = dotnetRef;
        if (this._bound) return;
        this._bound = true;

        document.addEventListener('pointerdown', function (e) {
            var handle = e.target.closest('.di-rh-e, .di-rh-s, .di-rh-se');
            if (!handle) return;

            e.preventDefault();
            e.stopPropagation();

            var wrap = handle.closest('.di-resizable');
            if (!wrap) return;

            var doW = handle.classList.contains('di-rh-e') || handle.classList.contains('di-rh-se');
            var doH = handle.classList.contains('di-rh-s') || handle.classList.contains('di-rh-se');

            var blockIdx = parseInt(wrap.dataset.blockIdx);
            var startX   = e.clientX;
            var startY   = e.clientY;
            var startW   = wrap.offsetWidth;
            var startH   = wrap.offsetHeight;
            var parentW  = wrap.parentElement.offsetWidth;

            // Block native drag (images etc.) while resizing
            function blockDrag(ev) { ev.preventDefault(); }
            document.addEventListener('dragstart', blockDrag, true);

            // Capture pointer so events keep firing even outside the handle
            try { handle.setPointerCapture(e.pointerId); } catch(_) {}
            document.body.style.userSelect = 'none';
            document.body.style.cursor = doW && doH ? 'nwse-resize' : doW ? 'ew-resize' : 'ns-resize';

            function cleanup() {
                handle.removeEventListener('pointermove', onMove);
                handle.removeEventListener('pointerup',   onUp);
                handle.removeEventListener('pointercancel', onUp);
                document.removeEventListener('dragstart', blockDrag, true);
                document.body.style.userSelect = '';
                document.body.style.cursor = '';
            }

            function onMove(ev) {
                if (doW) {
                    var pct = Math.max(10, Math.min(100,
                        Math.round((startW + ev.clientX - startX) / parentW * 100)));
                    wrap.style.width = pct + '%';
                    var bar = wrap.closest('.di-block');
                    var lbl = bar && bar.querySelector('.di-width-lbl');
                    if (lbl) lbl.textContent = pct + '%';
                }
                if (doH) {
                    var newH = Math.max(60, startH + ev.clientY - startY);
                    wrap.style.minHeight = newH + 'px';
                }
            }

            function onUp() {
                cleanup();
                var wPct = Math.round(wrap.offsetWidth / parentW * 100);
                var hPx  = Math.round(parseFloat(wrap.style.minHeight) || 0);
                if (window.diResize._dotnet)
                    window.diResize._dotnet.invokeMethodAsync('SetBlockSize', blockIdx, wPct, hPx);
            }

            handle.addEventListener('pointermove',   onMove);
            handle.addEventListener('pointerup',     onUp);
            handle.addEventListener('pointercancel', onUp);  // fired when browser cancels capture
        });
    }
};

// ── TipTap Editor ─────────────────────────────────────────────────────────────
window.tiptapEditor = {
    _editor: null,
    _mods:   null,

    async _load() {
        if (this._mods) return this._mods;
        const base = 'https://esm.sh/@tiptap/';
        const [
            { Editor },
            { default: StarterKit },
            { Table },
            { TableRow },
            { TableHeader },
            { TableCell },
            { Image: Img },
            { TextAlign },
        ] = await Promise.all([
            import(base + 'core@2'),
            import(base + 'starter-kit@2'),
            import(base + 'extension-table@2'),
            import(base + 'extension-table-row@2'),
            import(base + 'extension-table-header@2'),
            import(base + 'extension-table-cell@2'),
            import(base + 'extension-image@2'),
            import(base + 'extension-text-align@2'),
        ]);
        // Extend TableCell & TableHeader to preserve inline style attributes
        const styleAttr = {
            style: {
                default: null,
                parseHTML: el => el.getAttribute('style'),
                renderHTML: a => a.style ? { style: a.style } : {}
            }
        };
        const ExtTableCell   = TableCell.extend({   addAttributes() { return { ...this.parent?.(), ...styleAttr }; } });
        const ExtTableHeader = TableHeader.extend({ addAttributes() { return { ...this.parent?.(), ...styleAttr }; } });

        this._mods = { Editor, StarterKit, Table, TableRow, TableHeader, TableCell, Img, ExtTableCell, ExtTableHeader, TextAlign };
        return this._mods;
    },

    async init(elementId) {
        const el = document.getElementById(elementId);
        if (!el) return;
        if (this._editor) { this._editor.destroy(); this._editor = null; }

        const { Editor, StarterKit, Table, TableRow, Img, ExtTableCell, ExtTableHeader, TextAlign } = await this._load();

        // Image extension with SE resize handle
        const ResizableImg = Img.extend({
            addAttributes() {
                return {
                    ...this.parent?.(),
                    width: {
                        default: null,
                        parseHTML: el => el.style.width || el.getAttribute('width') || null,
                        renderHTML: attrs => attrs.width ? { style: `width:${attrs.width};height:auto;max-width:100%;` } : {}
                    }
                };
            },
            addNodeView() {
                return ({ node, updateAttributes }) => {
                    const wrap = document.createElement('span');
                    wrap.style.cssText = 'display:inline-block;position:relative;line-height:0;max-width:100%;';

                    const img = document.createElement('img');
                    img.src = node.attrs.src || '';
                    img.alt = node.attrs.alt || '';
                    img.style.cssText = 'display:block;max-width:100%;height:auto;';
                    if (node.attrs.width) img.style.width = node.attrs.width;

                    const handle = document.createElement('div');
                    handle.style.cssText = 'position:absolute;bottom:0;right:0;width:14px;height:14px;' +
                        'cursor:nwse-resize;border-right:3px solid #2563eb;border-bottom:3px solid #2563eb;' +
                        'border-radius:0 0 4px 0;opacity:0;transition:opacity .15s;z-index:10;';

                    wrap.addEventListener('mouseenter', () => handle.style.opacity = '1');
                    wrap.addEventListener('mouseleave', () => handle.style.opacity = '0');

                    handle.addEventListener('pointerdown', (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        handle.setPointerCapture(e.pointerId);

                        const startX = e.clientX;
                        const startW = img.offsetWidth || parseInt(node.attrs.width) || img.naturalWidth || 200;
                        document.body.style.userSelect = 'none';
                        document.body.style.cursor = 'nwse-resize';

                        const onMove = (ev) => {
                            img.style.width = Math.max(30, startW + ev.clientX - startX) + 'px';
                        };
                        const cleanup = () => {
                            handle.removeEventListener('pointermove', onMove);
                            handle.removeEventListener('pointerup', cleanup);
                            handle.removeEventListener('pointercancel', cleanup);
                            document.body.style.userSelect = '';
                            document.body.style.cursor = '';
                            updateAttributes({ width: img.style.width });
                        };

                        handle.addEventListener('pointermove', onMove);
                        handle.addEventListener('pointerup', cleanup);
                        handle.addEventListener('pointercancel', cleanup);
                    });

                    wrap.appendChild(img);
                    wrap.appendChild(handle);

                    return {
                        dom: wrap,
                        update(updated) {
                            if (updated.type !== node.type) return false;
                            if (img.src !== (updated.attrs.src || '')) img.src = updated.attrs.src || '';
                            if (updated.attrs.width) img.style.width = updated.attrs.width;
                            return true;
                        }
                    };
                };
            }
        }).configure({ inline: true, allowBase64: true });

        this._editor = new Editor({
            element: el,
            extensions: [
                StarterKit,
                Table.configure({ resizable: true }),
                TableRow,
                ExtTableHeader,
                ExtTableCell,
                ResizableImg,
                TextAlign.configure({ types: ['heading', 'paragraph'] }),
            ],
            content: [
                /* ── Title header table ─────────────────────────── */
                /* width:1% + white-space:nowrap = shrink-to-content */
                '<table><tbody>',
                '<tr>',
                  '<th rowspan="3" style="width:1%;white-space:nowrap;text-align:center;vertical-align:middle;">TITLE</th>',
                  '<td rowspan="3" style="text-align:center;font-weight:700;font-size:15px;vertical-align:middle;padding:14px 60px;"></td>',
                  '<th style="width:1%;white-space:nowrap;">Dept</th>',
                  '<td style="width:1%;white-space:nowrap;min-width:80px;"></td>',
                '</tr>',
                '<tr><th style="width:1%;white-space:nowrap;">Date</th><td></td></tr>',
                '<tr><th style="width:1%;white-space:nowrap;">Marker</th><td></td></tr>',
                '</tbody></table>',
                /* ── Sections I – IV ───────────────────────────── */
                '<h2>I. Purpose.</h2><p>- </p>' + '<p></p>'.repeat(5),
                '<h2>II. Content.</h2>' + '<p></p>'.repeat(10),
                '<h2>III. Result</h2>' + '<p></p>'.repeat(5),
                '<h2>IV. Decision</h2>' + '<p></p>'.repeat(5),
            ].join(''),
            editorProps: {
                attributes: { spellcheck: 'false' },
                handlePaste: (_view, event) => {
                    const items = event.clipboardData?.items;
                    if (!items) return false;
                    for (const item of items) {
                        if (!item.type.startsWith('image/')) continue;
                        const file = item.getAsFile();
                        if (!file) continue;
                        event.preventDefault();
                        const reader = new FileReader();
                        reader.onload = (ev) => {
                            window.tiptapEditor._editor
                                ?.chain().focus().setImage({ src: ev.target.result }).run();
                        };
                        reader.readAsDataURL(file);
                        return true;
                    }
                    return false;
                }
            },
        });
    },

    insertHTML(html) {
        this._editor?.chain().focus().insertContent(html).run();
    },

    insertImage(src, alt) {
        this._editor?.chain().focus().setImage({ src, alt: alt || '' }).run();
    },

    getHTML() { return this._editor?.getHTML() ?? ''; },

    getHTMLAndImages() {
        const html = this._editor?.getHTML() ?? '';
        const images = [];
        let idx = 0;
        const processed = html.replace(/src="data:([^;]+);base64,([^"]*)"/g, (_, mediaType, base64) => {
            const ext = mediaType.replace('image/', '').replace('jpeg', 'jpg').replace('svg+xml', 'svg');
            const slug = `di-img-${idx++}.${ext}`;
            images.push({ slug, base64 });
            return `src="di-img://${slug}"`;
        });
        return JSON.stringify({ html: processed, images });
    },

    setContentWithImages(html, imageMapJson) {
        let map = {};
        try { map = JSON.parse(imageMapJson); } catch (_) {}
        const restored = html.replace(/src="di-img:\/\/([^"]+)"/g, (match, slug) =>
            map[slug] ? `src="${map[slug]}"` : match
        );
        this._editor?.commands.setContent(restored || '', false);
    },

    setContent(html) { this._editor?.commands.setContent(html || '', false); },

    cmd(name) {
        if (!this._editor) return;
        const c = this._editor.chain().focus();
        switch (name) {
            case 'bold':         c.toggleBold().run();                    break;
            case 'italic':       c.toggleItalic().run();                  break;
            case 'strike':       c.toggleStrike().run();                  break;
            case 'h1':           c.toggleHeading({ level: 1 }).run();     break;
            case 'h2':           c.toggleHeading({ level: 2 }).run();     break;
            case 'h3':           c.toggleHeading({ level: 3 }).run();     break;
            case 'bullet':       c.toggleBulletList().run();              break;
            case 'ordered':      c.toggleOrderedList().run();             break;
            case 'hr':           c.setHorizontalRule().run();             break;
            case 'undo':         c.undo().run();                          break;
            case 'redo':         c.redo().run();                          break;
            case 'alignLeft':    c.setTextAlign('left').run();                break;
            case 'alignCenter':  c.setTextAlign('center').run();              break;
            case 'alignRight':   c.setTextAlign('right').run();               break;
            case 'alignJustify': c.setTextAlign('justify').run();             break;
            case 'insertTable':  c.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(); break;
            case 'addRowAfter':  c.addRowAfter().run();                   break;
            case 'addRowBefore': c.addRowBefore().run();                  break;
            case 'deleteRow':    c.deleteRow().run();                     break;
            case 'addColAfter':  c.addColumnAfter().run();                break;
            case 'addColBefore': c.addColumnBefore().run();               break;
            case 'deleteCol':    c.deleteColumn().run();                  break;
        }
    },

    initResize(elementId) {
        const STORAGE_KEY = 'di-editor-size';
        const editor = document.getElementById(elementId);
        if (!editor) return;
        const bg   = editor.closest('.di-tiptap-bg');
        const area = editor.closest('.di-tiptap-area');
        if (!bg || !area) return;

        area.querySelector('.di-editor-resize-handle')?.remove();
        area.style.position = 'relative';

        // Restore saved size
        try {
            const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
            if (saved?.w && saved?.h) {
                bg.style.flex   = 'none';
                bg.style.width  = saved.w + 'px';
                bg.style.height = saved.h + 'px';
            }
        } catch (_) {}

        const handle = document.createElement('div');
        handle.className = 'di-editor-resize-handle';
        area.appendChild(handle);

        handle.addEventListener('dblclick', () => {
            bg.style.flex   = '1';
            bg.style.width  = '';
            bg.style.height = '';
            localStorage.removeItem(STORAGE_KEY);
        });

        handle.addEventListener('pointerdown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            handle.setPointerCapture(e.pointerId);

            const startX = e.clientX, startY = e.clientY;
            const startW = bg.offsetWidth, startH = bg.offsetHeight;

            bg.style.flex   = 'none';
            bg.style.width  = startW + 'px';
            bg.style.height = startH + 'px';

            document.body.style.userSelect = 'none';
            document.body.style.cursor = 'nwse-resize';

            const onMove = (ev) => {
                bg.style.width  = Math.max(300, startW + ev.clientX - startX) + 'px';
                bg.style.height = Math.max(150, startH + ev.clientY - startY) + 'px';
            };
            const cleanup = () => {
                handle.removeEventListener('pointermove', onMove);
                handle.removeEventListener('pointerup', cleanup);
                handle.removeEventListener('pointercancel', cleanup);
                document.body.style.userSelect = '';
                document.body.style.cursor = '';
                // Save final size
                try {
                    localStorage.setItem(STORAGE_KEY, JSON.stringify({
                        w: Math.round(bg.offsetWidth),
                        h: Math.round(bg.offsetHeight)
                    }));
                } catch (_) {}
            };

            handle.addEventListener('pointermove', onMove);
            handle.addEventListener('pointerup', cleanup);
            handle.addEventListener('pointercancel', cleanup);
        });
    },

    destroy() {
        if (this._editor) { this._editor.destroy(); this._editor = null; }
    }
};

// ── Full-page layout helper ───────────────────────────────────────────────────
window.diFullPage = {
    enter: function () {
        var a = document.querySelector('article.content');
        if (!a) return;
        a._origStyle = a.getAttribute('style') || '';
        a.style.cssText = 'padding:0 !important; display:flex; flex-direction:column; flex:1; min-height:0; overflow:hidden;';
    },
    leave: function () {
        var a = document.querySelector('article.content');
        if (!a) return;
        a.style.cssText = a._origStyle || '';
    }
};

// ── Paste Image Handler ───────────────────────────────────────────────────────
window.pasteImageHandler = {
    _dotnetRef: null,
    _captureActive: false,
    _docPasteListener: null,

    // Called once on firstRender — just stores the dotnet ref and wires the document listener
    init: function (dotnetRef) {
        this._dotnetRef = dotnetRef;

        this._docPasteListener = (e) => {
            if (!this._captureActive || !this._dotnetRef) return; // only active during capture mode
            const items = e.clipboardData && e.clipboardData.items;
            if (!items) return;

            for (let i = 0; i < items.length; i++) {
                const item = items[i];
                if (!item.type.startsWith('image/')) continue;
                const file = item.getAsFile();
                if (!file) continue;

                e.preventDefault();
                this._captureActive = false;

                const reader = new FileReader();
                reader.onload = ev => {
                    // Flatten onto white background (kills transparency from Excel "No fill" cells)
                    const img = new Image();
                    img.onload = () => {
                        const canvas = document.createElement('canvas');
                        canvas.width  = img.naturalWidth  || img.width;
                        canvas.height = img.naturalHeight || img.height;
                        const ctx = canvas.getContext('2d');
                        ctx.fillStyle = '#ffffff';
                        ctx.fillRect(0, 0, canvas.width, canvas.height);
                        ctx.drawImage(img, 0, 0);
                        const flat = canvas.toDataURL('image/png');
                        this._dotnetRef.invokeMethodAsync('OnImagePasted', 'image/png', flat);
                    };
                    img.onerror = () => {
                        // Fallback: send original if canvas path fails
                        this._dotnetRef.invokeMethodAsync('OnImagePasted', item.type, ev.target.result);
                    };
                    img.src = ev.target.result;
                };
                reader.readAsDataURL(file);
                return;
            }
            // No image — text pastes into focused element normally (capture stays open)
        };

        document.addEventListener('paste', this._docPasteListener);
    },

    // Enable capture mode (called from Blazor @onclick — no focus trick needed for document-level paste)
    openCapture: function () { this._captureActive = true; },

    // Disable capture mode (Cancel button / ESC)
    cancelCapture: function () { this._captureActive = false; },

    // Cleanup on dispose
    cancelPasteCapture: function () {
        this._captureActive = false;
        if (this._docPasteListener) {
            document.removeEventListener('paste', this._docPasteListener);
            this._docPasteListener = null;
        }
        this._dotnetRef = null;
    }
};

// ── Backup File Drop/Paste Handler (any file type, not just images) ──────────
window.backupFileHandler = {
    _dotnetRef: null,
    _dropZone: null,
    _docPasteListener: null,
    _captureActive: false,

    init: function (dotnetRef, dropZoneId) {
        this._dotnetRef = dotnetRef;
        this._dropZone  = document.getElementById(dropZoneId);
        if (this._dropZone) {
            this._dropZone.addEventListener('dragover', this._onDragOver);
            this._dropZone.addEventListener('dragleave', this._onDragLeave);
            this._dropZone.addEventListener('drop', (e) => this._onDrop(e));
        }

        this._docPasteListener = (e) => {
            if (!this._captureActive || !this._dotnetRef) return;
            const items = e.clipboardData && e.clipboardData.items;
            if (!items) return;
            let handled = false;
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                if (it.kind !== 'file') continue;
                const file = it.getAsFile();
                if (!file) continue;
                handled = true;
                this._sendFile(file);
            }
            if (handled) e.preventDefault();
        };
        document.addEventListener('paste', this._docPasteListener);
    },

    _onDragOver: function (e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
        e.currentTarget.classList.add('bf-drop-active');
    },

    _onDragLeave: function (e) {
        e.currentTarget.classList.remove('bf-drop-active');
    },

    _onDrop: function (e) {
        e.preventDefault();
        if (this._dropZone) this._dropZone.classList.remove('bf-drop-active');
        if (!this._dotnetRef) return;
        const files = e.dataTransfer && e.dataTransfer.files;
        if (!files || files.length === 0) return;
        for (let i = 0; i < files.length; i++) this._sendFile(files[i]);
    },

    _sendFile: function (file) {
        const MAX = 100 * 1024 * 1024; // 100 MB
        if (file.size > MAX) {
            console.warn('File too large, skipped:', file.name, file.size);
            return;
        }
        const reader = new FileReader();
        reader.onload = (ev) => {
            this._dotnetRef.invokeMethodAsync(
                'OnBackupFileDropped',
                file.name || 'file',
                file.type || 'application/octet-stream',
                ev.target.result);
        };
        reader.readAsDataURL(file);
    },

    enableCapture: function () { this._captureActive = true;  },
    disableCapture: function () { this._captureActive = false; },

    dispose: function () {
        this._captureActive = false;
        if (this._docPasteListener) {
            document.removeEventListener('paste', this._docPasteListener);
            this._docPasteListener = null;
        }
        this._dropZone = null;
        this._dotnetRef = null;
    }
};
