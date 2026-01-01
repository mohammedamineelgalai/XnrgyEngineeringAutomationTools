using System;

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services;

public class SimpleLogger
{
	private readonly string _category;

	public SimpleLogger(string category)
	{
		_category = category;
	}

	public void LogInformation(string message)
	{
		Logger.Info("[" + _category + "] " + message);
	}

	public void LogDebug(string message)
	{
		Logger.Debug("[" + _category + "] " + message);
	}

	public void LogWarning(string message)
	{
		Logger.Warning("[" + _category + "] " + message);
	}

	public void LogError(string message)
	{
		Logger.Error("[" + _category + "] " + message);
	}

	public void LogError(Exception ex, string message)
	{
		Logger.Error("[" + _category + "] " + message + " - " + ex.Message);
		Logger.Error("[" + _category + "] StackTrace: " + ex.StackTrace);
	}
}

