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
