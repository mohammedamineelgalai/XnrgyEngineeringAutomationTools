#nullable enable
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Models;

public class LifecycleDefinitionItem : INotifyPropertyChanged
{
    public long Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public ObservableCollection<LifecycleStateItem> States { get; set; } = new();

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}


