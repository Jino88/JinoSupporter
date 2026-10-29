// Team memo: Ctrl+V of an image (screenshot, copied picture) inside a [data-tm-paste] zone attaches it.
// The pasted files are handed to the zone's hidden <InputFile class="tm-paste-input"> and a change event
// is raised, so the upload goes through Blazor's normal streamed InputFile path instead of base64 interop.
(function () {
    if (window.__tmPasteBound) return;
    window.__tmPasteBound = true;

    function stamp() {
        const d = new Date(), p = n => String(n).padStart(2, '0');
        return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
    }

    document.addEventListener('paste', function (e) {
        const zone = e.target && e.target.closest ? e.target.closest('[data-tm-paste]') : null;
        if (!zone || !e.clipboardData) return;
        const input = zone.querySelector('input[type=file].tm-paste-input');
        if (!input) return;

        const images = [];
        for (const item of e.clipboardData.items) {
            if (item.kind === 'file' && item.type.startsWith('image/')) {
                const f = item.getAsFile();
                if (f) images.push(f);
            }
        }
        if (images.length === 0) return;

        // A copied Excel range also carries a picture of itself; keep that a text paste.
        const text = e.clipboardData.getData('text/plain');
        if (text && text.trim().length > 0) return;

        e.preventDefault();
        const dt = new DataTransfer();
        images.forEach((f, i) => {
            const ext = (f.type.split('/')[1] || 'png').replace('jpeg', 'jpg');
            const name = f.name && f.name !== 'image.png'
                ? f.name
                : `pasted-${stamp()}${images.length > 1 ? '-' + (i + 1) : ''}.${ext}`;
            dt.items.add(new File([f], name, { type: f.type }));
        });
        input.files = dt.files;
        input.dispatchEvent(new Event('change', { bubbles: true }));
    });
})();
