#nullable enable
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Models;

public class LifecycleStateItem : INotifyPropertyChanged
{
    public long Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public long LifecycleDefinitionId { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}


