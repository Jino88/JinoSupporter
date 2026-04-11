using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace JinoSupporter.App.Modules.Schedule;

public partial class ScheduleView : UserControl
{
    // ── 팔레트 ────────────────────────────────────────────────────────────────
    private static readonly string[] ColorPalette =
    [
        "#3B82F6", "#10B981", "#F59E0B", "#EF4444",
        "#8B5CF6", "#EC4899", "#06B6D4", "#84CC16"
    ];

    private readonly ScheduleRepository _repo = new();

    private int    _currentYear;
    private int    _currentMonth;
    private string _selectedColor = ColorPalette[0];
    private readonly HashSet<string> _activeTagFilters = new(StringComparer.OrdinalIgnoreCase);

    // CalendarGrid의 Row 0 = 요일 헤더, Row 1-6 = 주(week) 행들
    // 각 주 행은 2개의 Grid Row를 차지: [날짜 번호 행, 이벤트 바 행]
    // 즉 CalendarGrid rows = 1 (header) + 6*2 (weeks) = 13

    public ScheduleView()
    {
        InitializeComponent();
        BuildColorPicker();
        _currentYear  = DateTime.Today.Year;
        _currentMonth = DateTime.Today.Month;
        Loaded += (_, _) => Refresh();
    }

    // ── 색상 선택기 ──────────────────────────────────────────────────────────
    private void BuildColorPicker()
    {
        ColorPickerPanel.Children.Clear();
        foreach (string hex in ColorPalette)
        {
            string h = hex; // capture
            var circle = new Border
            {
                Width        = 22,
                Height       = 22,
                CornerRadius = new CornerRadius(11),
                Background   = BrushFromHex(hex),
                Margin       = new Thickness(0, 0, 6, 0),
                Cursor       = Cursors.Hand,
                ToolTip      = hex
            };
            circle.MouseLeftButtonUp += (_, _) => SelectColor(h);
            ColorPickerPanel.Children.Add(circle);
        }
        SelectColor(_selectedColor);
    }

    private void SelectColor(string hex)
    {
        _selectedColor = hex;
        foreach (Border b in ColorPickerPanel.Children.OfType<Border>())
        {
            string bHex = b.ToolTip as string ?? string.Empty;
            b.BorderThickness = new Thickness(string.Equals(bHex, hex, StringComparison.OrdinalIgnoreCase) ? 2.5 : 0);
            b.BorderBrush     = Brushes.Black;
        }
    }

    // ── 네비게이션 ──────────────────────────────────────────────────────────
    private void PrevMonthButton_Click(object sender, RoutedEventArgs e)
    {
        _currentMonth--;
        if (_currentMonth < 1) { _currentMonth = 12; _currentYear--; }
        Refresh();
    }

    private void NextMonthButton_Click(object sender, RoutedEventArgs e)
    {
        _currentMonth++;
        if (_currentMonth > 12) { _currentMonth = 1; _currentYear++; }
        Refresh();
    }

    private void TodayButton_Click(object sender, RoutedEventArgs e)
    {
        _currentYear  = DateTime.Today.Year;
        _currentMonth = DateTime.Today.Month;
        Refresh();
    }

    // ── 메인 갱신 ────────────────────────────────────────────────────────────
    private void Refresh()
    {
        MonthYearLabel.Text = new DateTime(_currentYear, _currentMonth, 1).ToString("yyyy년 MM월");
        RebuildTagFilterChips();
        RebuildCalendar();
        RebuildUpcoming();
    }

    // ── 달력 구성 ────────────────────────────────────────────────────────────
    private void RebuildCalendar()
    {
        CalendarGrid.Children.Clear();
        CalendarGrid.RowDefinitions.Clear();

        // 칼럼 정의는 XAML에 이미 7개 있으므로 행만 추가
        // Row 0: 요일 헤더
        // Row 1..12: 주 (날짜 행 + 이벤트 바 행) × 6주
        CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(32) }); // header
        for (int w = 0; w < 6; w++)
        {
            CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(36) }); // date numbers
            CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // events
        }

        // 요일 헤더
        string[] dayNames = ["일", "월", "화", "수", "목", "금", "토"];
        for (int d = 0; d < 7; d++)
        {
            var tb = new TextBlock
            {
                Text              = dayNames[d],
                FontSize          = 12,
                FontWeight        = FontWeights.SemiBold,
                Foreground        = d == 0 ? new SolidColorBrush(Color.FromRgb(239, 68, 68))
                                  : d == 6 ? new SolidColorBrush(Color.FromRgb(59, 130, 246))
                                  : new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            Grid.SetRow(tb, 0);
            Grid.SetColumn(tb, d);
            CalendarGrid.Children.Add(tb);
        }

        // 이번 달 날짜 계산
        var firstDay  = new DateOnly(_currentYear, _currentMonth, 1);
        int daysInMonth = DateTime.DaysInMonth(_currentYear, _currentMonth);
        int startCol  = (int)firstDay.DayOfWeek; // 0=일

        // 달력에 표시할 첫 날 (이전 달 포함)
        DateOnly calStart = firstDay.AddDays(-startCol);
        DateOnly calEnd   = calStart.AddDays(42 - 1); // 6주

        // 해당 범위의 스케쥴 로드 (태그 필터 적용)
        List<ScheduleItem> schedules = _repo.GetSchedulesInRange(calStart, calEnd,
            _activeTagFilters.Count > 0 ? _activeTagFilters.ToList() : null);

        DateOnly today = DateOnly.FromDateTime(DateTime.Today);

        // 날짜 셀 채우기
        for (int cell = 0; cell < 42; cell++)
        {
            int week = cell / 7;
            int col  = cell % 7;
            DateOnly cellDate = calStart.AddDays(cell);
            bool isCurrentMonth = cellDate.Month == _currentMonth;
            bool isToday        = cellDate == today;

            int dateRow  = 1 + week * 2;   // 날짜 번호 행
            int eventRow = 2 + week * 2;   // 이벤트 바 행

            // 날짜 셀 배경
            var cellBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                Margin       = new Thickness(1),
                Background   = isToday
                    ? new SolidColorBrush(Color.FromRgb(59, 130, 246))
                    : new SolidColorBrush(Colors.Transparent),
                Cursor = Cursors.Hand
            };
            var dayNum = new TextBlock
            {
                Text      = cellDate.Day.ToString(),
                FontSize  = 12,
                FontWeight = isToday ? FontWeights.Bold : FontWeights.Normal,
                Foreground = isToday
                    ? Brushes.White
                    : isCurrentMonth
                        ? (col == 0 ? new SolidColorBrush(Color.FromRgb(239, 68, 68))
                                    : col == 6 ? new SolidColorBrush(Color.FromRgb(59, 130, 246))
                                    : new SolidColorBrush(Color.FromRgb(15, 23, 42)))
                        : new SolidColorBrush(Color.FromRgb(203, 213, 225)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            cellBorder.Child = dayNum;

            // 마우스오버 효과
            DateOnly captured = cellDate;
            cellBorder.MouseEnter += (_, _) =>
            {
                if (!isToday) cellBorder.Background = new SolidColorBrush(Color.FromRgb(239, 246, 255));
            };
            cellBorder.MouseLeave += (_, _) =>
            {
                cellBorder.Background = isToday
                    ? new SolidColorBrush(Color.FromRgb(59, 130, 246))
                    : Brushes.Transparent;
            };
            cellBorder.MouseLeftButtonUp += (_, _) => OnDateCellClicked(captured);

            Grid.SetRow(cellBorder, dateRow);
            Grid.SetColumn(cellBorder, col);
            CalendarGrid.Children.Add(cellBorder);
        }

        // ── 이벤트 바 배치 ─────────────────────────────────────────────────
        // 각 주(week)에 대해 해당 주와 겹치는 스케쥴 바를 그린다
        for (int week = 0; week < 6; week++)
        {
            DateOnly weekStart = calStart.AddDays(week * 7);
            DateOnly weekEnd   = weekStart.AddDays(6);
            int eventRow = 2 + week * 2;

            // 이 주와 겹치는 스케쥴
            var weekSchedules = schedules
                .Where(s => s.StartDate <= weekEnd && s.EndDate >= weekStart)
                .ToList();

            // 이벤트 바 높이: 16px + 2px 간격
            const int barHeight = 16;
            const int barMargin = 2;

            // 이 주 안에서의 시작/끝 컬럼 계산 후 Grid에 추가
            int lane = 0;
            foreach (ScheduleItem sched in weekSchedules)
            {
                int startColInWeek = Math.Max(0, sched.StartDate.DayNumber - weekStart.DayNumber);
                int endColInWeek   = Math.Min(6, sched.EndDate.DayNumber - weekStart.DayNumber);
                int span = endColInWeek - startColInWeek + 1;

                // 스케쥴 바
                string tooltip = $"{sched.Title}\n{sched.StartDate:yyyy-MM-dd} ~ {sched.EndDate:yyyy-MM-dd}  {sched.TimeDisplay}";
                if (!string.IsNullOrWhiteSpace(sched.Description)) tooltip += $"\n{sched.Description}";
                if (sched.Tags.Count > 0) tooltip += $"\n🏷 {sched.TagsDisplay}";

                var bar = new Border
                {
                    Background    = BrushFromHex(sched.Color, 220),
                    CornerRadius  = new CornerRadius(4),
                    Margin        = new Thickness(2, barMargin + lane * (barHeight + barMargin), 2, 0),
                    Height        = barHeight,
                    VerticalAlignment = VerticalAlignment.Top,
                    Cursor        = Cursors.Hand,
                    ToolTip       = tooltip
                };

                var barText = new TextBlock
                {
                    Text          = sched.IsAllDay ? sched.Title : $"{sched.StartTime!.Value.ToString("HH:mm")} {sched.Title}",
                    FontSize      = 10,
                    Foreground    = Brushes.White,
                    VerticalAlignment   = VerticalAlignment.Center,
                    TextTrimming  = TextTrimming.CharacterEllipsis,
                    Margin        = new Thickness(4, 0, 4, 0)
                };
                bar.Child = barText;

                // 삭제 컨텍스트 메뉴
                long schedId = sched.Id;
                var ctx = new ContextMenu();
                var deleteItem = new MenuItem { Header = "스케쥴 삭제" };
                deleteItem.Click += (_, _) =>
                {
                    _repo.DeleteSchedule(schedId);
                    Refresh();
                };
                ctx.Items.Add(deleteItem);
                bar.ContextMenu = ctx;

                Grid.SetRow(bar, eventRow);
                Grid.SetColumn(bar, startColInWeek);
                Grid.SetColumnSpan(bar, span);
                CalendarGrid.Children.Add(bar);

                lane++;
            }
        }
    }

    // ── 날짜 셀 클릭 시 폼에 날짜 채우기 ────────────────────────────────────
    private void OnDateCellClicked(DateOnly date)
    {
        StartDatePicker.SelectedDate = date.ToDateTime(TimeOnly.MinValue);
        if (EndDatePicker.SelectedDate == null ||
            DateOnly.FromDateTime(EndDatePicker.SelectedDate.Value) < date)
        {
            EndDatePicker.SelectedDate = date.ToDateTime(TimeOnly.MinValue);
        }
        NewTitleBox.Focus();
    }

    // ── 신규 스케쥴 추가 ─────────────────────────────────────────────────────
    private void AddScheduleButton_Click(object sender, RoutedEventArgs e)
    {
        string title = NewTitleBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            MessageBox.Show("제목을 입력해 주세요.", "스케쥴 추가", MessageBoxButton.OK, MessageBoxImage.Warning);
            NewTitleBox.Focus();
            return;
        }

        if (StartDatePicker.SelectedDate is null)
        {
            MessageBox.Show("시작일을 선택해 주세요.", "스케쥴 추가", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        DateOnly start = DateOnly.FromDateTime(StartDatePicker.SelectedDate.Value);
        DateOnly end   = EndDatePicker.SelectedDate.HasValue
            ? DateOnly.FromDateTime(EndDatePicker.SelectedDate.Value)
            : start;

        if (end < start)
        {
            MessageBox.Show("종료일은 시작일 이후여야 합니다.", "스케쥴 추가", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string desc = NewDescBox.Text.Trim();
        List<string> tags = NewTagsBox.Text
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();

        TimeOnly? startTime = null, endTime = null;
        string stStr = StartTimeBox.Text.Trim();
        string etStr = EndTimeBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(stStr))
        {
            if (!TimeOnly.TryParse(stStr, out TimeOnly st))
            {
                MessageBox.Show("시작 시간 형식이 올바르지 않습니다.\nHH:mm 형식으로 입력해 주세요. (예: 09:30)",
                    "스케쥴 추가", MessageBoxButton.OK, MessageBoxImage.Warning);
                StartTimeBox.Focus();
                return;
            }
            startTime = st;
        }
        if (!string.IsNullOrWhiteSpace(etStr))
        {
            if (!TimeOnly.TryParse(etStr, out TimeOnly et))
            {
                MessageBox.Show("종료 시간 형식이 올바르지 않습니다.\nHH:mm 형식으로 입력해 주세요. (예: 17:00)",
                    "스케쥴 추가", MessageBoxButton.OK, MessageBoxImage.Warning);
                EndTimeBox.Focus();
                return;
            }
            endTime = et;
        }

        _repo.AddSchedule(title, desc, start, end, _selectedColor, tags, startTime, endTime);

        NewTitleBox.Text = string.Empty;
        NewDescBox.Text  = string.Empty;
        NewTagsBox.Text  = string.Empty;
        StartTimeBox.Text = string.Empty;
        EndTimeBox.Text   = string.Empty;
        StartDatePicker.SelectedDate = null;
        EndDatePicker.SelectedDate   = null;

        // 추가한 달로 이동
        _currentYear  = start.Year;
        _currentMonth = start.Month;
        Refresh();
    }

    private void TagFilterClearButton_Click(object sender, RoutedEventArgs e)
    {
        _activeTagFilters.Clear();
        RebuildTagFilterChips();
        Refresh();
    }

    // ── 태그 필터 칩 렌더링 ──────────────────────────────────────────────────
    private void RebuildTagFilterChips()
    {
        TagFilterChipPanel.Children.Clear();

        // 전체 태그 칩 (토글)
        List<string> allTags = _repo.GetAllDistinctTags();
        foreach (string tag in allTags)
        {
            string captured = tag;
            bool active = _activeTagFilters.Contains(tag);

            var chip = new Border
            {
                Margin        = new Thickness(0, 0, 4, 4),
                Padding       = new Thickness(6, 2, 6, 2),
                CornerRadius  = new CornerRadius(10),
                Background    = active
                    ? new SolidColorBrush(Color.FromRgb(59, 130, 246))
                    : new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                Cursor        = Cursors.Hand
            };
            chip.Child = new TextBlock
            {
                Text       = tag,
                FontSize   = 10,
                Foreground = active ? Brushes.White
                                    : new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
            chip.MouseLeftButtonUp += (_, _) =>
            {
                if (_activeTagFilters.Contains(captured)) _activeTagFilters.Remove(captured);
                else _activeTagFilters.Add(captured);
                RebuildTagFilterChips();
                Refresh();
            };
            TagFilterChipPanel.Children.Add(chip);
        }
    }

    // ── 이번주 스케쥴 패널 ──────────────────────────────────────────────────
    private void RebuildUpcoming()
    {
        UpcomingPanel.Children.Clear();

        DateOnly today = DateOnly.FromDateTime(DateTime.Today);
        DateOnly until = today.AddDays(6);
        UpcomingRangeLabel.Text = $"{today:MM/dd} – {until:MM/dd}"
            + (_activeTagFilters.Count > 0 ? $"  [태그: {string.Join(", ", _activeTagFilters)}]" : string.Empty);

        List<ScheduleItem> items = _repo.GetUpcoming(7,
            _activeTagFilters.Count > 0 ? _activeTagFilters.ToList() : null);

        if (items.Count == 0)
        {
            UpcomingPanel.Children.Add(new TextBlock
            {
                Text       = "이번 주 등록된 스케쥴이 없습니다.",
                FontSize   = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                Margin     = new Thickness(0, 4, 0, 0)
            });
            return;
        }

        foreach (ScheduleItem item in items)
        {
            string dateRange = item.IsMultiDay
                ? $"{item.StartDate:MM/dd} – {item.EndDate:MM/dd}"
                : item.StartDate.ToString("MM/dd (ddd)");
            string timePart = item.IsAllDay ? "종일" : item.TimeDisplay;
            string dateRange2 = $"{dateRange}  {timePart}";

            var row = new Border
            {
                Margin       = new Thickness(0, 0, 0, 6),
                CornerRadius = new CornerRadius(8),
                Background   = BrushFromHex(item.Color, 30),
                BorderBrush  = BrushFromHex(item.Color, 120),
                BorderThickness = new Thickness(0, 0, 0, 0),
                Padding      = new Thickness(8, 6, 8, 6)
            };

            var inner = new Grid();
            inner.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
            inner.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var accent = new Border
            {
                Background   = BrushFromHex(item.Color),
                CornerRadius = new CornerRadius(2),
                Margin       = new Thickness(0, 0, 8, 0)
            };
            Grid.SetColumn(accent, 0);
            inner.Children.Add(accent);

            var textStack = new StackPanel { Margin = new Thickness(8, 0, 0, 0) };
            textStack.Children.Add(new TextBlock
            {
                Text       = item.Title,
                FontSize   = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(15, 23, 42)),
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            textStack.Children.Add(new TextBlock
            {
                Text       = dateRange2,
                FontSize   = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                Margin     = new Thickness(0, 2, 0, 0)
            });
            if (!string.IsNullOrWhiteSpace(item.Description))
            {
                textStack.Children.Add(new TextBlock
                {
                    Text         = item.Description,
                    FontSize     = 10,
                    Foreground   = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin       = new Thickness(0, 1, 0, 0)
                });
            }
            if (item.Tags.Count > 0)
            {
                textStack.Children.Add(new TextBlock
                {
                    Text         = "🏷 " + item.TagsDisplay,
                    FontSize     = 10,
                    Foreground   = new SolidColorBrush(Color.FromRgb(99, 102, 241)),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin       = new Thickness(0, 2, 0, 0)
                });
            }
            Grid.SetColumn(textStack, 1);
            inner.Children.Add(textStack);

            row.Child = inner;

            // 우클릭 삭제
            long schedId = item.Id;
            var ctx = new ContextMenu();
            var del = new MenuItem { Header = "스케쥴 삭제" };
            del.Click += (_, _) =>
            {
                _repo.DeleteSchedule(schedId);
                Refresh();
            };
            ctx.Items.Add(del);
            row.ContextMenu = ctx;

            UpcomingPanel.Children.Add(row);
        }
    }

    // ── 유틸 ────────────────────────────────────────────────────────────────
    private static SolidColorBrush BrushFromHex(string hex, byte alpha = 255)
    {
        try
        {
            hex = hex.TrimStart('#');
            byte r = Convert.ToByte(hex.Substring(0, 2), 16);
            byte g = Convert.ToByte(hex.Substring(2, 2), 16);
            byte b = Convert.ToByte(hex.Substring(4, 2), 16);
            return new SolidColorBrush(Color.FromArgb(alpha, r, g, b));
        }
        catch
        {
            return Brushes.Gray;
        }
    }
}
