using System;
using System.ComponentModel;
using System.Windows.Controls;
using Microsoft.Win32;

namespace GraphMaker;

/// <summary>
/// 모든 GraphMaker View UserControl의 공통 기반.
/// 아래 중복 코드를 한 곳에서 관리합니다:
///   - WebModuleSnapshotChanged 이벤트 / NotifyWebModuleSnapshotChanged()
///   - INotifyPropertyChanged / OnPropertyChanged()
///   - 공통 파일 다이얼로그 헬퍼 (TryBrowseFiles)
///   - 구분자 라디오 버튼 → 문자열 변환 (ResolveDelimiter)
/// </summary>
public abstract class GraphViewBase : UserControl, INotifyPropertyChanged
{
    // ───────────────────────────────────────────────────
    // Shell 연동: 스냅샷 변경 알림
    // ───────────────────────────────────────────────────

    public event Action? WebModuleSnapshotChanged;

    protected void NotifyWebModuleSnapshotChanged() => WebModuleSnapshotChanged?.Invoke();

    // ───────────────────────────────────────────────────
    // INotifyPropertyChanged
    // ───────────────────────────────────────────────────

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    // ───────────────────────────────────────────────────
    // 파일 선택 다이얼로그 헬퍼
    // ───────────────────────────────────────────────────

    /// <summary>
    /// OpenFileDialog를 열고 선택된 파일 경로 배열을 반환합니다.
    /// 사용자가 취소하면 false를 반환하고 files는 빈 배열입니다.
    /// </summary>
    protected static bool TryBrowseFiles(
        string title,
        string filter,
        bool multiselect,
        out string[] files)
    {
        var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = filter,
            Multiselect = multiselect
        };

        if (dialog.ShowDialog() == true)
        {
            files = dialog.FileNames;
            return true;
        }

        files = Array.Empty<string>();
        return false;
    }

    // ───────────────────────────────────────────────────
    // 구분자 헬퍼
    // ───────────────────────────────────────────────────

    /// <summary>
    /// Tab / Comma / Space 라디오 버튼의 IsChecked 상태를 읽어
    /// 해당 구분자 문자열을 반환합니다. 아무것도 체크되지 않으면 탭("\t")을 반환합니다.
    /// </summary>
    protected static string ResolveDelimiter(
        RadioButton? tabRadio,
        RadioButton? commaRadio,
        RadioButton? spaceRadio)
    {
        if (tabRadio?.IsChecked == true) return "\t";
        if (commaRadio?.IsChecked == true) return ",";
        if (spaceRadio?.IsChecked == true) return " ";
        return "\t";
    }
}
