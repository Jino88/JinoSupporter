import type { ViewerDefaults } from "../contract";
import { preferencesFromDefaults, type ViewerPreferences } from "../logic";

export function ViewerToolbar({
  preferences,
  defaults,
  onChange,
}: {
  preferences: ViewerPreferences;
  defaults: ViewerDefaults;
  onChange: (next: ViewerPreferences) => void;
}) {
  const setNumber = (key: keyof ViewerPreferences, rawValue: string) => {
    const value = Number(rawValue);
    onChange({ ...preferences, [key]: Number.isFinite(value) && value >= 0 ? value : 0 });
  };

  return (
    <div className="viewer-toolbar" aria-label="리포트 표시 필터">
      <label>
        <span>Minimum PPM</span>
        <input
          type="number"
          min="0"
          step="100"
          value={preferences.minimumPpm}
          onChange={(event) => setNumber("minimumPpm", event.currentTarget.value)}
        />
      </label>
      <fieldset>
        <legend>표시 기간</legend>
        <label><span>일</span><input type="number" min="0" step="1" value={preferences.dateColumnLimit} onChange={(event) => setNumber("dateColumnLimit", event.currentTarget.value)} /></label>
        <label><span>주</span><input type="number" min="0" step="1" value={preferences.weekColumnLimit} onChange={(event) => setNumber("weekColumnLimit", event.currentTarget.value)} /></label>
        <label><span>월</span><input type="number" min="0" step="1" value={preferences.monthColumnLimit} onChange={(event) => setNumber("monthColumnLimit", event.currentTarget.value)} /></label>
      </fieldset>
      <button type="button" className="secondary-button" onClick={() => onChange(preferencesFromDefaults(defaults))}>기본값</button>
      <button type="button" className="secondary-button" onClick={() => window.print()}>인쇄</button>
    </div>
  );
}
