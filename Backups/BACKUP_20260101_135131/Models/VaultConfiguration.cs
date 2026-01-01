using System.ComponentModel.DataAnnotations;

namespace XnrgyEngineeringAutomationTools.Models;

public class VaultConfiguration
{
	[Required]
	public string ServerName { get; set; } = string.Empty;

	[Required]
	public string VaultName { get; set; } = string.Empty;

	[Required]
	public string Username { get; set; } = string.Empty;

	[Required]
	public string Password { get; set; } = string.Empty;

	public string Domain { get; set; }

	public int ConnectionTimeoutSeconds { get; set; } = 30;

	public int RetryAttempts { get; set; } = 3;

	public int RetryDelayMs { get; set; } = 1000;
}
