namespace XnrgyEngineeringAutomationTools.Models;

public class ModuleInfo
{
	public string FullPath { get; set; }

	public string DisplayName { get; set; }

	public string ProjectNumber { get; set; }

	public string Reference { get; set; }

	public string Module { get; set; }

	public override string ToString()
	{
		return DisplayName;
	}
}
