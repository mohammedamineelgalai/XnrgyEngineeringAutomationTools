namespace XnrgyEngineeringAutomationTools.Models;

public class CategoryItem
{
    public long Id { get; set; }
    public string Name { get; set; } = string.Empty;
    
    public override string ToString() => Name;
}


