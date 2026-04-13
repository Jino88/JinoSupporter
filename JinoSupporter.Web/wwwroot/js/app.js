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

        // separator 위치 (빈 라벨 "") 인덱스 집합
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
                    // Send only dataUrl — C# side extracts base64 to avoid doubling payload size
                    this._dotnetRef.invokeMethodAsync('OnImagePasted', item.type, ev.target.result);
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
