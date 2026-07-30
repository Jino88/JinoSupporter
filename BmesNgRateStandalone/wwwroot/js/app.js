// ?? NG Rate Summary Chart ?????????????????????????????????????????????????????
// Returns the current page's document.title so the tab strip can label a tab with the
// page's own <PageTitle> rather than a name guessed from the URL.
window.appDocTitle = function () { return document.title || ''; };

// Reading document.title once after a render is not enough: Blazor's <PageTitle> writes it
// asynchronously, so a read taken right after navigating still returns the PREVIOUS page's
// title and the new tab gets mislabelled. Watch <head> for title changes and push each one
// to the layout instead.
window.appTitleWatcher = {
    _observer: null,

    start: function (dotNetRef) {
        this.stop();
        if (!dotNetRef || !document.head) return;

        const notify = () => {
            const title = (document.title || '').trim();
            if (!title) return;
            // The circuit may already be gone when a page is torn down.
            dotNetRef.invokeMethodAsync('OnDocumentTitleChanged', title).catch(() => { });
        };

        // subtree/childList so a replaced <title> element is caught, not just edited text.
        this._observer = new MutationObserver(notify);
        this._observer.observe(document.head, { childList: true, characterData: true, subtree: true });
        notify();   // seed the tab we landed on
    },

    stop: function () {
        if (this._observer) {
            this._observer.disconnect();
            this._observer = null;
        }
    }
};

window.ngRateLog = {
    _observers: {},

    scrollToBottom: function (elementId, markerId) {
        const el = document.getElementById(elementId);
        if (!el) {
            [25, 75, 150, 300, 600, 1000].forEach(delay =>
                setTimeout(() => window.ngRateLog.scrollToBottom(elementId, markerId), delay));
            return;
        }

        const scroll = () => {
            el.style.overflowAnchor = 'none';
            el.scrollTop = el.scrollHeight;
        };

        const schedule = () => {
            scroll();
            requestAnimationFrame(() => {
                scroll();
                requestAnimationFrame(scroll);
            });
            [0, 25, 75, 150, 300, 600, 1000, 1500].forEach(delay => setTimeout(scroll, delay));
        };

        const existing = window.ngRateLog._observers[elementId];
        if (!existing || existing.el !== el) {
            if (existing && existing.observer) existing.observer.disconnect();

            const observer = new MutationObserver(schedule);
            observer.observe(el, { childList: true, subtree: true, characterData: true });
            window.ngRateLog._observers[elementId] = { el, observer };
        }

        schedule();
    }
};

window.ngRateTableStyle = {
    _observer: null,
    _scheduled: false,

    start: function () {
        this.schedule();
        if (this._observer || !document.body) return;

        this._observer = new MutationObserver(() => this.schedule());
        this._observer.observe(document.body, { childList: true, subtree: true });
    },

    schedule: function () {
        if (this._scheduled) return;
        this._scheduled = true;
        requestAnimationFrame(() => {
            this._scheduled = false;
            this.apply();
        });
    },

    apply: function () {
        document.querySelectorAll('.pivot-table').forEach(table => this.applyTable(table));
    },

    applyTable: function (table) {
        const rows = Array.from(table.tBodies || []).flatMap(body => Array.from(body.rows || []));
        const maxByColumn = new Map();

        rows.forEach(row => {
            Array.from(row.cells || []).forEach(cell => cell.classList.remove('ppm-max-cell'));
        });

        rows.filter(row => this.isEligibleRow(row)).forEach(row => {
            Array.from(row.cells || []).forEach((cell, columnIndex) => {
                if (!this.isDataCell(cell)) return;

                const value = this.parseNumber(cell.textContent);
                if (value == null || value <= 0) return;

                const current = maxByColumn.get(columnIndex);
                if (!current || value > current.value) {
                    maxByColumn.set(columnIndex, { value, cells: [cell] });
                } else if (Math.abs(value - current.value) < 0.5) {
                    current.cells.push(cell);
                }
            });
        });

        maxByColumn.forEach(entry => {
            if (entry.cells.length === 0) return;
            entry.cells.forEach(cell => cell.classList.add('ppm-max-cell'));
        });
    },

    isEligibleRow: function (row) {
        const skip = [
            'total-row',
            'hier-group-row',
            'grp-overall-row',
            'grp-total-row',
            'wr-mid-row'
        ];
        return !skip.some(className => row.classList.contains(className));
    },

    isDataCell: function (cell) {
        if (!cell || cell.tagName !== 'TD') return false;
        const skip = ['label-td', 'sep-td', 'toggle-cell', 'row-hide-cell'];
        return !skip.some(className => cell.classList.contains(className));
    },

    parseNumber: function (text) {
        const match = (text || '').replace(/\([^)]*\)/g, '').match(/-?\d[\d,]*(?:\.\d+)?/);
        if (!match) return null;
        const value = Number(match[0].replace(/,/g, ''));
        return Number.isFinite(value) ? value : null;
    }
};

document.addEventListener('DOMContentLoaded', () => window.ngRateTableStyle.start());

window.bmesReportTableSizer = {
    _observer: null,
    _scheduled: false,
    _resizeAttached: false,

    start: function () {
        this.schedule();
        if (!this._resizeAttached) {
            this._resizeAttached = true;
            window.addEventListener('resize', () => this.schedule());
        }
        if (document.fonts && document.fonts.ready) {
            document.fonts.ready.then(() => this.schedule()).catch(() => {});
        }
        if (this._observer || !document.body) return;

        this._observer = new MutationObserver(() => this.schedule());
        this._observer.observe(document.body, { childList: true, subtree: true });
    },

    schedule: function () {
        if (this._scheduled) return;
        this._scheduled = true;
        requestAnimationFrame(() => {
            this._scheduled = false;
            this.apply();
        });
    },

    apply: function () {
        const root = document.querySelector('.bmes-report-tab-content');
        if (!root) return;

        root.querySelectorAll('.pivot-table.text-fit-table')
            .forEach(table => this.applyTable(table));
    },

    applyTable: function (table) {
        table.querySelectorAll('.bmes-report-number-cell')
            .forEach(cell => cell.classList.remove('bmes-report-number-cell'));

        const numericCells = Array.from(table.querySelectorAll('td'))
            .filter(cell => this.isNumericValueCell(cell));

        numericCells.forEach(cell => cell.classList.add('bmes-report-number-cell'));

        const width = this.measureMaxNumberCellWidth(table);
        if (!width) {
            table.style.removeProperty('--bmes-report-cell-width');
            return;
        }

        table.style.setProperty('--bmes-report-cell-width', `${width}px`);
    },

    measureMaxNumberCellWidth: function (table) {
        const clone = table.cloneNode(true);
        clone.removeAttribute('id');
        clone.style.setProperty('--bmes-report-cell-width', '');
        clone.style.setProperty('table-layout', 'auto', 'important');
        clone.style.setProperty('width', 'max-content', 'important');
        clone.style.setProperty('min-width', '0', 'important');
        clone.style.setProperty('max-width', 'none', 'important');

        clone.querySelectorAll('th,td').forEach(cell => {
            cell.classList.remove('bmes-report-number-cell');
            cell.style.setProperty('width', 'auto', 'important');
            cell.style.setProperty('min-width', '0', 'important');
            cell.style.setProperty('max-width', 'none', 'important');
        });

        const host = document.createElement('div');
        host.style.position = 'fixed';
        host.style.left = '-100000px';
        host.style.top = '0';
        host.style.visibility = 'hidden';
        host.style.pointerEvents = 'none';
        host.style.width = 'max-content';
        host.style.maxWidth = 'none';
        host.appendChild(clone);
        document.body.appendChild(host);

        try {
            let max = 0;
            clone.querySelectorAll('td').forEach(cell => {
                if (!this.isNumericValueCell(cell)) return;
                const width = Math.ceil(cell.getBoundingClientRect().width);
                if (width > max) max = width;
            });

            return max > 0 ? Math.max(48, max + 6) : 0;
        } finally {
            host.remove();
        }
    },

    isNumericValueCell: function (cell) {
        if (!cell || cell.tagName !== 'TD' || cell.colSpan !== 1) return false;
        const skip = [
            'label-td',
            'group-name-td',
            'sep-td',
            'toggle-cell',
            'row-hide-cell'
        ];
        if (skip.some(className => cell.classList.contains(className))) return false;
        return this.parseNumber(cell.textContent) != null;
    },

    parseNumber: function (text) {
        const normalized = (text || '').replace(/\([^)]*\)/g, '').trim();
        if (!normalized || normalized === '-') return null;

        const match = normalized.match(/-?\d[\d,]*(?:\.\d+)?/);
        if (!match) return null;

        const value = Number(match[0].replace(/,/g, ''));
        return Number.isFinite(value) ? value : null;
    }
};

document.addEventListener('DOMContentLoaded', () => window.bmesReportTableSizer.start());

window.ngRateImageCopy = {
    copyElementToClipboard: async function (elementId) {
        const source = document.getElementById(elementId);
        if (!source) throw new Error(`Copy target not found: ${elementId}`);
        if (!navigator.clipboard || !window.ClipboardItem) {
            throw new Error('This browser does not support image clipboard copy.');
        }

        if (document.fonts && document.fonts.ready) {
            await document.fonts.ready;
        }

        let blob = null;
        let renderError = null;

        try {
            const canvas = await this.renderToCanvas(source);
            blob = await this.canvasToBlob(canvas);
        } catch (ex) {
            renderError = ex;
        }

        if (!blob && source.querySelector('.pivot-table, table')) {
            try {
                const canvas = this.renderTablesToCanvas(source);
                blob = await this.canvasToBlob(canvas);
            } catch (ex) {
                renderError = renderError || ex;
            }
        }

        if (!blob) {
            throw renderError || new Error('Failed to render the table image.');
        }

        await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob })
        ]);
    },

    canvasToBlob: function (canvas) {
        return new Promise((resolve, reject) => {
            try {
                canvas.toBlob(blob => {
                    if (blob) resolve(blob);
                    else reject(new Error('Failed to render the table image.'));
                }, 'image/png');
            } catch (ex) {
                reject(ex);
            }
        });
    },

    renderToCanvas: async function (source) {
        const initialSize = this.measure(source);
        const host = document.createElement('div');
        host.style.position = 'fixed';
        host.style.left = '-100000px';
        host.style.top = '0';
        host.style.width = `${initialSize.width}px`;
        host.style.background = '#fff';
        host.style.pointerEvents = 'none';
        host.style.zIndex = '-1';

        const clone = source.cloneNode(true);
        clone.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
        host.appendChild(clone);
        document.body.appendChild(host);

        try {
            this.inlineStyles(source, clone);
            this.prepareClone(clone, initialSize.width);

            const width = Math.ceil(Math.max(initialSize.width, clone.scrollWidth, clone.getBoundingClientRect().width));
            const height = Math.ceil(Math.max(initialSize.height, clone.scrollHeight, clone.getBoundingClientRect().height));
            clone.style.width = `${width}px`;
            host.style.width = `${width}px`;

            const serialized = new XMLSerializer().serializeToString(clone);
            const svg = [
                `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">`,
                '<foreignObject width="100%" height="100%">',
                serialized,
                '</foreignObject>',
                '</svg>'
            ].join('');

            const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }));
            try {
                const img = await this.loadImage(url);
                const scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1));
                const canvas = document.createElement('canvas');
                canvas.width = Math.ceil(width * scale);
                canvas.height = Math.ceil(height * scale);
                canvas.style.width = `${width}px`;
                canvas.style.height = `${height}px`;

                const ctx = canvas.getContext('2d');
                ctx.setTransform(scale, 0, 0, scale, 0, 0);
                ctx.fillStyle = '#fff';
                ctx.fillRect(0, 0, width, height);
                ctx.drawImage(img, 0, 0, width, height);
                return canvas;
            } finally {
                URL.revokeObjectURL(url);
            }
        } finally {
            host.remove();
        }
    },

    measure: function (source) {
        const rect = source.getBoundingClientRect();
        let width = Math.ceil(Math.max(1, rect.width, source.offsetWidth || 0, source.scrollWidth || 0));
        let height = Math.ceil(Math.max(1, rect.height, source.offsetHeight || 0, source.scrollHeight || 0));

        source.querySelectorAll('.pivot-wrap, table').forEach(el => {
            const r = el.getBoundingClientRect();
            width = Math.max(width, Math.ceil(r.width), el.offsetWidth || 0, el.scrollWidth || 0);
            height = Math.max(height, Math.ceil(r.bottom - rect.top), el.offsetHeight || 0, el.scrollHeight || 0);
        });

        return { width, height };
    },

    inlineStyles: function (source, clone) {
        const sourceNodes = [source, ...source.querySelectorAll('*')];
        const cloneNodes = [clone, ...clone.querySelectorAll('*')];

        sourceNodes.forEach((src, index) => {
            const dst = cloneNodes[index];
            if (!dst) return;
            const computed = window.getComputedStyle(src);
            let cssText = '';
            for (let i = 0; i < computed.length; i++) {
                const prop = computed[i];
                cssText += `${prop}:${computed.getPropertyValue(prop)};`;
            }
            dst.style.cssText = cssText;
        });
    },

    prepareClone: function (clone, width) {
        clone.querySelectorAll('[data-copy-hidden="true"]').forEach(el => el.remove());
        clone.style.background = '#fff';
        clone.style.overflow = 'visible';
        clone.style.width = `${width}px`;
        clone.style.maxWidth = 'none';
        clone.style.boxShadow = 'none';

        clone.querySelectorAll('.pivot-wrap').forEach(el => {
            el.style.overflow = 'visible';
            el.style.width = 'auto';
            el.style.maxWidth = 'none';
        });

        clone.querySelectorAll('.pivot-table, table').forEach(el => {
            el.style.setProperty('table-layout', 'auto', 'important');
            el.style.setProperty('width', 'max-content', 'important');
            el.style.setProperty('min-width', '100%', 'important');
            el.style.setProperty('max-width', 'none', 'important');
        });

        this.prepareTableTextForCopy(clone);

        clone.querySelectorAll('th').forEach(el => {
            el.style.position = 'static';
            el.style.top = 'auto';
        });
    },

    prepareTableTextForCopy: function (root) {
        root.querySelectorAll('.pivot-table th, .pivot-table td, table th, table td').forEach(el => {
            el.style.overflow = 'visible';
            el.style.textOverflow = 'clip';
        });

        const valueCells = Array.from(root.querySelectorAll('.pivot-table th:not(.label-th):not(.sep-th), .pivot-table td:not(.label-td):not(.sep-td)'))
            .filter(el => (el.colSpan || 1) === 1);

        valueCells.forEach(el => {
            const style = window.getComputedStyle(el);
            el.style.boxSizing = 'border-box';
            el.style.paddingLeft = `${Math.max(this.cssPx(style.paddingLeft), 5)}px`;
            el.style.paddingRight = `${Math.max(this.cssPx(style.paddingRight), 5)}px`;
            el.style.maxWidth = 'none';
        });

        root.querySelectorAll('.ppm-cell-value, .ppm-delta').forEach(el => {
            el.style.lineHeight = '1.22';
            el.style.overflow = 'visible';
        });

        this.applyCopyCellWidths(valueCells);
    },

    applyCopyCellWidths: function (cells) {
        if (!cells || !cells.length) return;

        const measurer = document.createElement('span');
        measurer.style.position = 'fixed';
        measurer.style.left = '-100000px';
        measurer.style.top = '0';
        measurer.style.visibility = 'hidden';
        measurer.style.pointerEvents = 'none';
        measurer.style.whiteSpace = 'nowrap';
        document.body.appendChild(measurer);

        try {
            cells.forEach(cell => {
                const style = window.getComputedStyle(cell);
                const paddingLeft = Math.max(this.cssPx(style.paddingLeft), 5);
                const paddingRight = Math.max(this.cssPx(style.paddingRight), 5);
                const measuredText = this.measureCellTextWidth(cell, measurer);
                const currentWidth = Math.ceil(cell.getBoundingClientRect().width);
                const floor = cell.tagName === 'TD' ? 64 : 54;
                const width = Math.ceil(Math.max(currentWidth, measuredText + paddingLeft + paddingRight + 10, floor));

                cell.style.setProperty('box-sizing', 'border-box', 'important');
                cell.style.setProperty('width', `${width}px`, 'important');
                cell.style.setProperty('min-width', `${width}px`, 'important');
                cell.style.setProperty('max-width', 'none', 'important');
            });
        } finally {
            measurer.remove();
        }
    },

    measureCellTextWidth: function (cell, measurer) {
        const parts = Array.from(cell.querySelectorAll('.ppm-cell-value, .ppm-delta'));
        if (parts.length) {
            return parts.reduce((max, part) => {
                const text = this.copyText(part);
                if (!text) return max;
                return Math.max(max, this.measureTextWidth(text, window.getComputedStyle(part), measurer));
            }, 0);
        }

        const text = this.copyText(cell);
        return text ? this.measureTextWidth(text, window.getComputedStyle(cell), measurer) : 0;
    },

    measureTextWidth: function (text, style, measurer) {
        measurer.style.font = this.canvasFont(style);
        measurer.style.fontVariantNumeric = style.fontVariantNumeric || 'normal';
        measurer.style.letterSpacing = style.letterSpacing || 'normal';
        measurer.textContent = text;
        return Math.ceil(measurer.getBoundingClientRect().width);
    },

    renderTablesToCanvas: function (source) {
        const initialSize = this.measure(source);
        const host = document.createElement('div');
        host.style.position = 'fixed';
        host.style.left = '-100000px';
        host.style.top = '0';
        host.style.width = `${initialSize.width}px`;
        host.style.background = '#fff';
        host.style.pointerEvents = 'none';
        host.style.zIndex = '-1';

        const clone = source.cloneNode(true);
        host.appendChild(clone);
        document.body.appendChild(host);

        try {
            this.prepareClone(clone, initialSize.width);
            return this.drawTablesFromElement(clone);
        } finally {
            host.remove();
        }
    },

    drawTablesFromElement: function (root) {
        const items = [];

        root.querySelectorAll('.card-header').forEach(header => {
            const rect = header.getBoundingClientRect();
            if (this.isRenderable(header, rect)) {
                items.push({ type: 'header', el: header, rect });
            }
        });

        root.querySelectorAll('.pivot-table, table').forEach(table => {
            Array.from(table.rows || []).forEach(row => {
                Array.from(row.cells || []).forEach(cell => {
                    const rect = cell.getBoundingClientRect();
                    if (this.isRenderable(cell, rect)) {
                        items.push({ type: 'cell', el: cell, rect });
                    }
                });
            });
        });

        if (!items.length) {
            throw new Error('No copyable table was found.');
        }

        const bounds = items.reduce((acc, item) => ({
            left: Math.min(acc.left, item.rect.left),
            top: Math.min(acc.top, item.rect.top),
            right: Math.max(acc.right, item.rect.right),
            bottom: Math.max(acc.bottom, item.rect.bottom)
        }), {
            left: Number.POSITIVE_INFINITY,
            top: Number.POSITIVE_INFINITY,
            right: Number.NEGATIVE_INFINITY,
            bottom: Number.NEGATIVE_INFINITY
        });

        const padding = 4;
        const rowHeightScale = 1;
        const width = Math.ceil(bounds.right - bounds.left + padding * 2);
        const height = Math.ceil((bounds.bottom - bounds.top) * rowHeightScale + padding * 2);
        const scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1));
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(width * scale);
        canvas.height = Math.ceil(height * scale);
        canvas.style.width = `${width}px`;
        canvas.style.height = `${height}px`;

        const ctx = canvas.getContext('2d');
        ctx.setTransform(scale, 0, 0, scale, 0, 0);
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, width, height);

        items
            .sort((a, b) => (a.rect.top - b.rect.top) || (a.rect.left - b.rect.left))
            .forEach(item => {
                const x = item.rect.left - bounds.left + padding;
                const y = (item.rect.top - bounds.top) * rowHeightScale + padding;
                const w = item.rect.width;
                const h = item.rect.height * rowHeightScale;

                if (item.type === 'header') {
                    this.drawHeader(ctx, item.el, x, y, w, h);
                } else {
                    this.drawCell(ctx, item.el, x, y, w, h);
                }
            });

        return canvas;
    },

    isRenderable: function (el, rect) {
        if (!rect || rect.width <= 0 || rect.height <= 0) return false;
        const style = window.getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden';
    },

    drawHeader: function (ctx, header, x, y, w, h) {
        const style = window.getComputedStyle(header);
        ctx.fillStyle = this.canvasColor(style.backgroundColor, '#f8fafc');
        ctx.fillRect(x, y, w, h);
        this.drawBorder(ctx, x, y, w, h, style);

        const text = this.copyText(header);
        if (!text) return;

        this.drawText(ctx, text, x + 8, y, Math.max(0, w - 16), h, style, {
            align: 'left',
            baseline: 'middle'
        });
    },

    drawCell: function (ctx, cell, x, y, w, h) {
        const style = window.getComputedStyle(cell);
        ctx.fillStyle = this.canvasColor(style.backgroundColor, '#fff');
        ctx.fillRect(x, y, w, h);
        this.drawBorder(ctx, x, y, w, h, style);

        ctx.save();
        ctx.beginPath();
        ctx.rect(x + 1, y + 1, Math.max(0, w - 2), Math.max(0, h - 2));
        ctx.clip();

        const valueEl = cell.querySelector('.ppm-cell-value');
        const deltaEl = cell.querySelector('.ppm-delta');
        if (valueEl) {
            const valueStyle = window.getComputedStyle(valueEl);
            const deltaStyle = deltaEl ? window.getComputedStyle(deltaEl) : null;
            const valueText = this.copyText(valueEl);
            const deltaText = deltaEl ? this.copyText(deltaEl) : '';

            if (deltaText) {
                const valueLineHeight = this.canvasLineHeight(valueStyle);
                const deltaLineHeight = this.canvasLineHeight(deltaStyle);
                const gap = Math.max(1, this.cssPx(deltaStyle.marginTop));
                const blockHeight = valueLineHeight + gap + deltaLineHeight;
                const blockTop = y + Math.max(1, (h - blockHeight) / 2);

                this.drawText(ctx, valueText, x, blockTop, w, valueLineHeight, valueStyle, {
                    align: 'right',
                    baseline: 'middle'
                });
                this.drawText(ctx, deltaText, x, blockTop + valueLineHeight + gap, w, deltaLineHeight, deltaStyle, {
                    align: 'right',
                    baseline: 'middle'
                });
            } else {
                this.drawText(ctx, valueText, x, y, w, h, valueStyle, {
                    align: style.textAlign,
                    baseline: 'middle'
                });
            }
        } else {
            const text = this.copyText(cell);
            if (text) {
                this.drawText(ctx, text, x, y, w, h, style, {
                    align: style.textAlign,
                    baseline: 'middle'
                });
            }
        }

        ctx.restore();
    },

    drawBorder: function (ctx, x, y, w, h, style) {
        const color = this.canvasColor(style.borderTopColor, '#9ca3af');
        ctx.save();
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1, this.cssPx(style.borderTopWidth) || 1);
        ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, w - 1), Math.max(0, h - 1));
        ctx.restore();
    },

    drawText: function (ctx, text, x, y, w, h, style, options) {
        const paddingLeft = this.cssPx(style.paddingLeft);
        const paddingRight = this.cssPx(style.paddingRight);
        const align = (options && options.align) || style.textAlign || 'left';
        const baseline = (options && options.baseline) || 'middle';
        const maxWidth = Math.max(0, w - paddingLeft - paddingRight);
        if (!text || maxWidth <= 0) return;

        ctx.save();
        ctx.font = this.canvasFont(style);
        ctx.fillStyle = this.canvasColor(style.color, '#111827');
        ctx.textBaseline = baseline;

        let textX = x + paddingLeft;
        ctx.textAlign = 'left';
        if (align === 'center') {
            textX = x + w / 2;
            ctx.textAlign = 'center';
        } else if (align === 'right' || align === 'end') {
            textX = x + w - paddingRight;
            ctx.textAlign = 'right';
        }

        let textY = y + h / 2;
        if (baseline === 'top') {
            textY = y;
        } else if (baseline === 'bottom') {
            textY = y + h;
        }

        ctx.fillText(this.fitText(ctx, text, maxWidth), textX, textY);
        ctx.restore();
    },

    fitText: function (ctx, text, maxWidth) {
        if (ctx.measureText(text).width <= maxWidth) return text;
        if (maxWidth <= ctx.measureText('...').width) return '';

        let lo = 0;
        let hi = text.length;
        while (lo < hi) {
            const mid = Math.ceil((lo + hi) / 2);
            if (ctx.measureText(text.slice(0, mid) + '...').width <= maxWidth) lo = mid;
            else hi = mid - 1;
        }
        return text.slice(0, lo) + '...';
    },

    copyText: function (el) {
        const clone = el.cloneNode(true);
        clone.querySelectorAll('[data-copy-hidden="true"], button, script, style').forEach(node => node.remove());
        return (clone.textContent || '').replace(/\s+/g, ' ').trim();
    },

    canvasFont: function (style) {
        return `${style.fontStyle || 'normal'} ${style.fontWeight || '400'} ${style.fontSize || '12px'} ${style.fontFamily || 'sans-serif'}`;
    },

    canvasLineHeight: function (style) {
        if (!style) return 12;

        const lineHeight = this.cssPx(style.lineHeight);
        if (lineHeight > 0) return lineHeight;

        const fontSize = this.cssPx(style.fontSize) || 12;
        return Math.ceil(fontSize * 1.2);
    },

    cssPx: function (value) {
        const n = Number.parseFloat(value || '0');
        return Number.isFinite(n) ? n : 0;
    },

    canvasColor: function (value, fallback) {
        if (!value || value === 'transparent') return fallback;
        const compact = value.replace(/\s+/g, '');
        if (compact === 'rgba(0,0,0,0)' || compact.endsWith(',0)')) return fallback;
        return value;
    },

    loadImage: function (url) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = () => reject(new Error('Failed to load rendered SVG image.'));
            img.src = url;
        });
    }
};

window.ngRateChart = {
    _instances: {},
    _valueLabelsPlugin: {
        id: 'ngRateSummaryValueLabels',
        afterDatasetsDraw: function (chart) {
            const options = chart.options.plugins.ngRateSummaryValueLabels || {};
            if (options.display === false) return;

            const ctx = chart.ctx;
            ctx.save();
            ctx.font = '600 10px sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillStyle = '#111827';

            chart.data.datasets.forEach((dataset, datasetIndex) => {
                if (dataset.type !== 'line' || !chart.isDatasetVisible(datasetIndex)) return;
                const meta = chart.getDatasetMeta(datasetIndex);

                meta.data.forEach((point, pointIndex) => {
                    const value = dataset.data[pointIndex];
                    if (value == null || !Number.isFinite(Number(value))) return;

                    const y = Math.max(chart.chartArea.top + 12, point.y - 8);
                    ctx.fillText(Number(value).toLocaleString(), point.x, y);
                });
            });

            ctx.restore();
        }
    },

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
            plugins: [this._valueLabelsPlugin],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    ngRateSummaryValueLabels: {
                        display: labels.filter(l => l !== '').length <= 14
                    },
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

// ?? Auto-reload on Blazor reconnect ??????????????????????????????????????????
// Blazor Server???쒕쾭 ?ъ떆????"Attempting to reconnect" 紐⑤떖???꾩슦怨?湲곕낯 8???ъ떆?????ㅽ뙣?섎㈃ 硫덉땅?덈떎.
// ?곕━???ъ뿰寃?紐⑤떖???⑤뒗 利됱떆 二쇨린?곸쑝濡??쒕쾭瑜??묓븯怨? ?묐떟???ㅻ㈃ ?섏씠吏瑜??섎뱶 由щ줈?쒗빐??// ?ъ슜?먭? 吏곸젒 ?덈줈怨좎묠?섏? ?딆븘????鍮뚮뱶媛 ?먮룞 諛섏쁺?섍쾶 ?⑸땲??
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

// ?? NG Rate By-Group Line Chart ???????????????????????????????????????????????
window.ngRateGroupChart = {
    _instances: {},
    _palette: [
        '#4f46e5', '#2563eb', '#14b8a6', '#f97316', '#ef4444',
        '#8b5cf6', '#0ea5e9', '#f59e0b', '#10b981', '#ec4899',
        '#64748b', '#a855f7', '#22c55e', '#eab308', '#06b6d4'
    ],
    _valueLabelsPlugin: {
        id: 'ngRateValueLabels',
        afterDatasetsDraw: function (chart) {
            const options = chart.options.plugins.ngRateValueLabels || {};
            if (options.display === false) return;

            const ctx = chart.ctx;
            ctx.save();
            ctx.font = '600 10px sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillStyle = '#111827';

            chart.data.datasets.forEach((dataset, datasetIndex) => {
                if (!chart.isDatasetVisible(datasetIndex)) return;
                const meta = chart.getDatasetMeta(datasetIndex);

                meta.data.forEach((point, pointIndex) => {
                    const value = dataset.data[pointIndex];
                    if (value == null || !Number.isFinite(Number(value))) return;

                    const y = Math.max(chart.chartArea.top + 12, point.y - 8);
                    ctx.fillText(Number(value).toLocaleString(), point.x, y);
                });
            });

            ctx.restore();
        }
    },

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
                backgroundColor: color + '18',
                borderWidth: 2,
                pointBackgroundColor: color,
                pointBorderColor: color,
                pointRadius: 3,
                pointHoverRadius: 5,
                spanGaps: false, // don't bridge across separators / missing points
                tension: 0.4,
            };
        });

        this._instances[canvasId] = new Chart(canvas, {
            type: 'line',
            data: { labels: labels, datasets: datasets },
            plugins: [this._valueLabelsPlugin],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    ngRateValueLabels: {
                        display: datasets.length <= 2 && labels.filter(l => l !== '').length <= 14
                    },
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 10,
                            padding: 8,
                            usePointStyle: true,
                            font: { size: 10 },
                            color: '#111827'
                        }
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
                        grid: { color: ctx => sepIdx.has(ctx.index) ? '#cbd5e1' : '#edf2f7' },
                        border: { color: '#cbd5e1' }
                    },
                    y: {
                        ticks: { font: { size: 10 }, callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v },
                        grid: { color: '#e5e7eb' },
                        border: { display: false },
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

window.graphMaker = {
    _instances: {},
    _palette: [
        '#2563eb', '#16a34a', '#dc2626', '#9333ea', '#ea580c',
        '#0891b2', '#be123c', '#4f46e5', '#65a30d', '#0f766e'
    ],

    render: function (canvasId, payload) {
        this.destroy(canvasId);
        const canvas = document.getElementById(canvasId);
        if (!canvas || !payload) return;
        const existing = window.Chart ? Chart.getChart(canvas) : null;
        if (existing) existing.destroy();
        payload = this.normalizePayload(payload);

        if (payload.kind === 'HeatMap') {
            this.renderHeatMap(canvas, payload);
            return;
        }

        const kind = payload.kind || 'Line';
        const isScatter = kind === 'Scatter' || kind === 'ProcessTrend' || kind === 'NormalDistribution';
        const datasets = [];

        (payload.series || []).forEach((s, i) => {
            const color = this._palette[i % this._palette.length];
            if (s.points) {
                datasets.push({
                    type: 'scatter',
                    label: s.name,
                    data: s.points,
                    borderColor: s.color || color,
                    backgroundColor: (s.color || color) + '33',
                    pointRadius: s.pointRadius ?? (s.isLimit || kind === 'NormalDistribution' ? 0 : 3),
                    showLine: !!s.showLine || kind === 'NormalDistribution',
                    borderWidth: s.isLimit ? 1.5 : 2,
                    borderDash: s.dashed ? [6, 4] : undefined,
                    _graphMakerLimit: !!s.isLimit,
                    _graphMakerLabelValue: s.labelValue,
                    tension: 0.2,
                });
            } else {
                const lineColor = s.color || color;
                datasets.push({
                    type: 'line',
                    label: s.name,
                    data: s.data,
                    borderColor: lineColor,
                    backgroundColor: lineColor + '22',
                    pointRadius: s.pointRadius ?? (kind === 'Line' ? 2 : 3),
                    showLine: kind !== 'NoXMultiY',
                    borderWidth: 2,
                    borderDash: s.dashed ? [6, 4] : undefined,
                    tension: payload.lineMode === 'Smoothing' ? 0.35 : 0,
                });
            }
        });

        (payload.limits || []).filter(l => l.value != null).forEach(l => {
            datasets.push({
                type: 'line',
                label: l.name,
                data: (payload.labels || []).map(() => l.value),
                borderColor: l.color,
                borderWidth: 1.5,
                borderDash: [6, 4],
                pointRadius: 0,
                tension: 0,
            });
        });

        this._instances[canvasId] = new Chart(canvas, {
            type: isScatter ? 'scatter' : 'line',
            data: {
                labels: payload.labels || [],
                datasets: datasets,
            },
            plugins: [this.limitLabelPlugin, this.categoryLabelPlugin],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        bottom: payload.useCategoryXAxis ? 96 : 0
                    }
                },
                parsing: isScatter ? false : true,
                interaction: { mode: 'nearest', intersect: false },
                plugins: {
                    graphMakerStats: payload.statGroups && payload.statGroups.length ? [] : (payload.stats || []),
                    graphMakerCategoryLabels: {
                        enabled: payload.useCategoryXAxis,
                        labels: payload.labels || [],
                        position: payload.scatterLabelPosition
                    },
                    title: { display: true, text: payload.title || 'Graph Maker' },
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 12, padding: 10, font: { size: 11 } }
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => {
                                if (ctx.raw && typeof ctx.raw === 'object') {
                                    return ` ${ctx.dataset.label}: (${this.format(ctx.raw.x)}, ${this.format(ctx.raw.y)})`;
                                }
                                return ` ${ctx.dataset.label}: ${this.format(ctx.parsed.y)}`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        type: isScatter ? 'linear' : 'category',
                        min: payload.useCategoryXAxis ? (payload.scatterLabelPosition === 'BetweenTicks' ? 0 : -0.5) : undefined,
                        max: payload.useCategoryXAxis ? (payload.scatterLabelPosition === 'BetweenTicks' ? (payload.labels || []).length : (payload.labels || []).length - 0.5) : undefined,
                        ticks: {
                            stepSize: payload.useCategoryXAxis ? 1 : undefined,
                            autoSkip: !payload.useCategoryXAxis,
                            maxRotation: 35,
                            font: { size: 10 },
                            callback: function (value) {
                                if (!payload.useCategoryXAxis) return value;
                                return '';
                            }
                        },
                        grid: { color: '#e2e8f0' }
                    },
                    y: {
                        ticks: { font: { size: 10 } },
                        grid: { color: '#e2e8f0' }
                    }
                }
            }
        });
    },

    normalizePayload: function (payload) {
        return {
            kind: payload.kind ?? payload.Kind,
            title: payload.title ?? payload.Title,
            labels: payload.labels ?? payload.Labels ?? [],
            useCategoryXAxis: payload.useCategoryXAxis ?? payload.UseCategoryXAxis,
            scatterLabelPosition: payload.scatterLabelPosition ?? payload.ScatterLabelPosition,
            lineMode: payload.lineMode ?? payload.LineMode,
            series: (payload.series ?? payload.Series ?? []).map(s => ({
                name: s.name ?? s.Name,
                data: s.data ?? s.Data,
                points: (s.points ?? s.Points)?.map(p => ({
                    x: p.x ?? p.X,
                    y: p.y ?? p.Y
                })),
                isLimit: s.isLimit ?? s.IsLimit,
                color: s.color ?? s.Color,
                showLine: s.showLine ?? s.ShowLine,
                dashed: s.dashed ?? s.Dashed,
                labelValue: s.labelValue ?? s.LabelValue,
                pointRadius: s.pointRadius ?? s.PointRadius
            })),
            limits: (payload.limits ?? payload.Limits ?? []).map(l => ({
                name: l.name ?? l.Name,
                value: l.value ?? l.Value,
                color: l.color ?? l.Color
            })),
            xLabels: payload.xLabels ?? payload.XLabels ?? [],
            yLabels: payload.yLabels ?? payload.YLabels ?? [],
            matrix: payload.matrix ?? payload.Matrix ?? [],
            matrixLabels: payload.matrixLabels ?? payload.MatrixLabels ?? [],
            statGroups: (payload.statGroups ?? payload.StatGroups ?? []).map(g => ({
                x: g.x ?? g.X,
                label: g.label ?? g.Label,
                stats: (g.stats ?? g.Stats ?? []).map(s => ({
                    name: s.name ?? s.Name,
                    value: s.value ?? s.Value
                }))
            })),
            stats: payload.stats ?? payload.Stats ?? []
        };
    },

    limitLabelPlugin: {
        id: 'graphMakerLimitLabels',
        afterDatasetsDraw: function (chart) {
            const ctx = chart.ctx;
            const yScale = chart.scales.y;
            const area = chart.chartArea;
            if (!yScale || !area) return;

            ctx.save();
            ctx.font = '11px sans-serif';
            ctx.fillStyle = '#111827';
            ctx.textAlign = 'right';
            ctx.textBaseline = 'middle';

            chart.data.datasets.forEach(ds => {
                if (!ds._graphMakerLimit || ds._graphMakerLabelValue == null) return;
                const y = yScale.getPixelForValue(ds._graphMakerLabelValue);
                if (!Number.isFinite(y)) return;
                const label = `${ds.label} ${window.graphMaker.format(ds._graphMakerLabelValue)}`;
                ctx.fillText(label, area.right - 6, y - 8);
            });

            const stats = chart.options.plugins.graphMakerStats || [];
            if (stats.length) {
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                ctx.fillStyle = 'rgba(255,255,255,.88)';
                const statLines = stats.map(s => {
                    const name = s.name ?? s.Name;
                    const value = s.value ?? s.Value;
                    return `${name}: ${window.graphMaker.format(value)}`;
                });
                const boxW = Math.max(128, Math.min(360, Math.ceil(Math.max(...statLines.map(line => ctx.measureText(line).width))) + 24));
                const boxH = 8 + stats.length * 18;
                ctx.fillRect(area.left + 8, area.top + 8, boxW, boxH);
                ctx.strokeStyle = '#cbd5e1';
                ctx.strokeRect(area.left + 8, area.top + 8, boxW, boxH);
                ctx.fillStyle = '#111827';
                statLines.forEach((line, i) => ctx.fillText(line, area.left + 16, area.top + 14 + i * 18));
            }

            ctx.restore();
        }
    },

    categoryLabelPlugin: {
        id: 'graphMakerCategoryLabels',
        afterDraw: function (chart) {
            const opts = chart.options.plugins.graphMakerCategoryLabels;
            if (!opts || !opts.enabled) return;

            const xScale = chart.scales.x;
            const area = chart.chartArea;
            if (!xScale || !area) return;

            const ctx = chart.ctx;
            ctx.save();
            ctx.font = '10px sans-serif';
            ctx.fillStyle = '#334155';
            ctx.textAlign = 'right';
            ctx.textBaseline = 'middle';

            const offset = opts.position === 'BetweenTicks' ? 0.5 : 0;
            opts.labels.forEach((label, i) => {
                const x = xScale.getPixelForValue(i + offset);
                if (!Number.isFinite(x)) return;
                ctx.save();
                ctx.translate(x, area.bottom + 44);
                ctx.rotate(-Math.PI / 4);
                ctx.fillText(window.graphMaker.truncateLabel(String(label)), 0, 0);
                ctx.restore();
            });

            ctx.restore();
        }
    },

    renderHeatMap: function (canvas, payload) {
        const ctx = canvas.getContext('2d');
        const container = canvas.parentElement;
        const width = Math.max(320, container ? container.clientWidth - 18 : 800);
        const height = Math.max(320, container ? container.clientHeight - 18 : 500);
        canvas.width = width * window.devicePixelRatio;
        canvas.height = height * window.devicePixelRatio;
        canvas.style.width = width + 'px';
        canvas.style.height = height + 'px';
        ctx.setTransform(window.devicePixelRatio, 0, 0, window.devicePixelRatio, 0, 0);
        ctx.clearRect(0, 0, width, height);

        const xLabels = payload.xLabels || [];
        const yLabels = payload.yLabels || [];
        const matrix = payload.matrix || [];
        const matrixLabels = payload.matrixLabels || [];
        const left = 110, top = 44, right = 24, bottom = 80;
        const plotW = Math.max(1, width - left - right);
        const plotH = Math.max(1, height - top - bottom);
        const cellW = plotW / Math.max(1, xLabels.length);
        const cellH = plotH / Math.max(1, yLabels.length);
        const values = matrix.flat().filter(v => typeof v === 'number');
        const min = values.length ? Math.min(...values) : 0;
        const max = values.length ? Math.max(...values) : 1;

        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
        ctx.fillStyle = '#1f2937';
        ctx.font = '600 14px sans-serif';
        ctx.fillText(payload.title || 'Heatmap', left, 24);

        const stats = matrixLabels.length ? [] : (payload.stats || []);
        if (stats.length) {
            ctx.font = '11px sans-serif';
            const statLines = stats.map(s => `${s.name ?? s.Name}: ${this.format(s.value ?? s.Value)}`);
            const boxW = Math.max(128, Math.min(360, Math.ceil(Math.max(...statLines.map(line => ctx.measureText(line).width))) + 24));
            const boxH = 8 + statLines.length * 18;
            const boxX = Math.max(left + 160, width - right - boxW);
            ctx.fillStyle = 'rgba(255,255,255,.9)';
            ctx.fillRect(boxX, 10, boxW, boxH);
            ctx.strokeStyle = '#cbd5e1';
            ctx.strokeRect(boxX, 10, boxW, boxH);
            ctx.fillStyle = '#111827';
            ctx.textAlign = 'left';
            ctx.textBaseline = 'top';
            statLines.forEach((line, i) => ctx.fillText(line, boxX + 8, 16 + i * 18));
        }

        for (let y = 0; y < yLabels.length; y++) {
            for (let x = 0; x < xLabels.length; x++) {
                const value = matrix[y] ? matrix[y][x] : null;
                ctx.fillStyle = value == null ? '#f8fafc' : this.heatColor((value - min) / Math.max(0.000001, max - min));
                ctx.fillRect(left + x * cellW, top + y * cellH, Math.ceil(cellW), Math.ceil(cellH));
                ctx.strokeStyle = '#ffffff';
                ctx.strokeRect(left + x * cellW, top + y * cellH, cellW, cellH);
                const labelText = matrixLabels[y] ? matrixLabels[y][x] : null;
                if (labelText && cellW > 52 && cellH > 48) {
                    ctx.fillStyle = '#111827';
                    ctx.font = '10px sans-serif';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    const lines = String(labelText).split('\n');
                    const startY = top + y * cellH + cellH / 2 - (lines.length - 1) * 6;
                    lines.forEach((line, lineIndex) => {
                        ctx.fillText(line, left + x * cellW + cellW / 2, startY + lineIndex * 12);
                    });
                } else if (value != null && cellW > 38 && cellH > 18) {
                    ctx.fillStyle = '#111827';
                    ctx.font = '10px sans-serif';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillText(this.format(value), left + x * cellW + cellW / 2, top + y * cellH + cellH / 2);
                }
            }
        }

        ctx.fillStyle = '#475569';
        ctx.font = '10px sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        yLabels.forEach((label, i) => ctx.fillText(String(label).slice(0, 18), left - 8, top + i * cellH + cellH / 2));

        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        xLabels.forEach((label, i) => {
            ctx.save();
            ctx.translate(left + i * cellW + cellW / 2, top + plotH + 8);
            ctx.rotate(-Math.PI / 4);
            ctx.fillText(String(label).slice(0, 18), 0, 0);
            ctx.restore();
        });

        this._instances[canvas.id] = { destroy: () => ctx.clearRect(0, 0, width, height) };
    },

    heatColor: function (t) {
        t = Math.max(0, Math.min(1, t || 0));
        const r = Math.round(37 + t * 218);
        const g = Math.round(99 + (1 - Math.abs(t - 0.5) * 2) * 120);
        const b = Math.round(235 - t * 180);
        return `rgb(${r},${g},${b})`;
    },

    format: function (value) {
        if (value == null || Number.isNaN(value)) return '-';
        if (Math.abs(value) >= 1000) return Number(value).toLocaleString(undefined, { maximumFractionDigits: 1 });
        return Number(value).toLocaleString(undefined, { maximumFractionDigits: 4 });
    },

    truncateLabel: function (label) {
        label = String(label || '');
        return label.length > 24 ? label.slice(0, 21) + '...' : label;
    },

    destroy: function (canvasId) {
        if (this._instances[canvasId]) {
            this._instances[canvasId].destroy();
            delete this._instances[canvasId];
        }
        const canvas = document.getElementById(canvasId);
        const existing = canvas && window.Chart ? Chart.getChart(canvas) : null;
        if (existing) existing.destroy();
    }
};

// ?? File Download ?????????????????????????????????????????????????????????????
window.graphMakerPaste = {
    _dotnet: null,
    _bound: false,

    init: function (dotnetRef) {
        this._dotnet = dotnetRef;
        if (this._bound) return;
        this._bound = true;
        document.addEventListener('paste', this._handlePaste);
    },

    destroy: function () {
        if (this._bound) {
            document.removeEventListener('paste', this._handlePaste);
        }
        this._bound = false;
        this._dotnet = null;
    },

    _handlePaste: function (e) {
        const sheet = document.getElementById('graphMakerSheet');
        if (!sheet || !window.graphMakerPaste._dotnet) return;

        const active = document.activeElement;
        const target = e.target;
        const isSheetPaste = sheet.contains(target) || active === sheet || sheet.contains(active);
        if (!isSheetPaste) return;

        const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
        if (!text || !text.trim()) return;

        e.preventDefault();
        window.graphMakerPaste._dotnet.invokeMethodAsync('PasteExcelData', text);
    }
};

window.downloadBase64File = function (filename, base64, contentType) {
    const link = document.createElement('a');
    link.href     = `data:${contentType};base64,${base64}`;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

window.downloadFileFromStream = async function (filename, contentStreamReference, contentType) {
    const arrayBuffer = await contentStreamReference.arrayBuffer();
    const blob = new Blob([arrayBuffer], { type: contentType || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
};

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

// ?? Data Inference Block Resize ??????????????????????????????????????????????
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

// ?? TipTap Editor ?????????????????????????????????????????????????????????????
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
                /* ?? Title header table ??????????????????????????? */
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
                /* ?? Sections I ??IV ????????????????????????????? */
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

// ?? Full-page layout helper ???????????????????????????????????????????????????
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

window.didbLazy = {
    _dotnet: null,
    _el: null,
    _onScroll: null,
    _pending: false,

    init: function (dotnetRef) {
        this.dispose();
        this._dotnet = dotnetRef;
        this._el = document.querySelector('.didb-table-wrap');
        if (!this._el) return;

        this._onScroll = () => {
            if (!this._dotnet || !this._el || this._pending) return;
            const remaining = this._el.scrollHeight - this._el.scrollTop - this._el.clientHeight;
            if (remaining > 700) return;

            this._pending = true;
            Promise.resolve(this._dotnet.invokeMethodAsync('LoadMoreDatasets'))
                .finally(() => window.setTimeout(() => { this._pending = false; }, 80));
        };

        this._el.addEventListener('scroll', this._onScroll, { passive: true });
        window.setTimeout(this._onScroll, 0);
    },

    dispose: function () {
        if (this._el && this._onScroll)
            this._el.removeEventListener('scroll', this._onScroll);
        this._dotnet = null;
        this._el = null;
        this._onScroll = null;
        this._pending = false;
    }
};

// ?? Paste Image Handler ???????????????????????????????????????????????????????
window.pasteImageHandler = {
    _dotnetRef: null,
    _captureActive: false,
    _docPasteListener: null,

    // Called once on firstRender ??just stores the dotnet ref and wires the document listener
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
            // No image ??text pastes into focused element normally (capture stays open)
        };

        document.addEventListener('paste', this._docPasteListener);
    },

    // Enable capture mode (called from Blazor @onclick ??no focus trick needed for document-level paste)
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

// ?? Backup File Drop/Paste Handler (any file type, not just images) ??????????
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

window.jinoDailyTest = window.jinoDailyTest || {};
window.jinoDailyTest.openHtml = function (title, html) {
    const win = window.open('', '_blank');
    if (!win) return false;

    win.document.open();
    win.document.write(html || '');
    win.document.close();
    try {
        win.document.title = title || 'Daily Test Data Analysis';
        win.focus();
    } catch (_) { }
    return true;
};
