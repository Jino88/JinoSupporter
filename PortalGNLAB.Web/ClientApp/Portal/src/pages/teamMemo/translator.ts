export interface TranslationSource {
    cached(lang: string, texts: string[]): Promise<Record<string, string>>;
    fetch(lang: string, texts: string[]): Promise<Record<string, string>>;
}
/** Generation checks stop a slow response for an old language replacing the current map. */
export class MemoTranslator {
    lang = "";
    busy = false;
    pending = 0;
    error = "";
    private generation = 0;
    private map: Record<string, string> = {};
    constructor(private source: TranslationSource, private redraw: () => void) { }
    t = (text: string | null | undefined): string => text ? (this.lang && Object.hasOwn(this.map, text) ? this.map[text] : text) : "";
    async refresh(lang: string, texts: (string | null | undefined)[]): Promise<void> {
        const generation = ++this.generation;
        if (this.lang !== lang)
            this.map = {};
        this.lang = lang;
        this.error = "";
        this.pending = 0;
        this.busy = false;
        this.redraw();
        if (!lang) {
            this.map = {};
            this.redraw();
            return;
        }
        const list = [...new Set(texts.filter((x): x is string => !!x?.trim()))];
        try {
            const cached = await this.source.cached(lang, list);
            if (generation !== this.generation)
                return;
            this.map = cached;
            const missing = list.filter(t => !Object.hasOwn(cached, t) && /\p{L}/u.test(t));
            this.pending = missing.length;
            this.busy = missing.length > 0;
            this.redraw();
            if (missing.length) {
                const fetched = await this.source.fetch(lang, missing);
                if (generation !== this.generation)
                    return;
                this.map = { ...this.map, ...fetched };
            }
        }
        catch (e) {
            if (generation === this.generation)
                this.error = "Translation failed: " + (e as Error).message;
        }
        finally {
            if (generation === this.generation) {
                this.pending = 0;
                this.busy = false;
                this.redraw();
            }
        }
    }
    dispose() { this.generation++; }
}
