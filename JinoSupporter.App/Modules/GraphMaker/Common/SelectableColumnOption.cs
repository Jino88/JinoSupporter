using System.ComponentModel;

namespace GraphMaker
{
    public sealed class SelectableColumnOption : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string ColumnName { get; set; } = string.Empty;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
