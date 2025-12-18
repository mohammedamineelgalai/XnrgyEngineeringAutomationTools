using System;
using System.IO;
using System.Text;

using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Services;

public static class Logger
{
	public enum LogLevel
	{
		TRACE,
		DEBUG,
		INFO,
		WARNING,
		ERROR,
		CRITICAL,
		FATAL
	}

	private static string _logFilePath;

	private static readonly object _lockObj = new();

	public static void Initialize()
	{
		try
		{
			string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
			string text = Path.Combine(baseDirectory, "Logs");
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			string text2 = DateTime.Now.ToString("yyyyMMdd_HHmmss");
			_logFilePath = Path.Combine(text, "VaultSDK_POC_" + text2 + ".log");
			Log("═══════════════════════════════════════════════════════");
			Log("  VAULT SDK POC - SESSION DEMARREE");
			Log($"  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
			Log("═══════════════════════════════════════════════════════");
			Log("Fichier log: " + _logFilePath);
		}
		catch (Exception ex)
		{
			Console.WriteLine("❌ Erreur initialisation logger: " + ex.Message);
		}
	}

	public static void Log(string message, LogLevel level = LogLevel.INFO)
	{
		if (string.IsNullOrEmpty(_logFilePath))
		{
			Initialize();
		}
		try
		{
			lock (_lockObj)
			{
				string text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
				string text2 = level.ToString().PadRight(7);
				string text3 = "[" + text + "] [" + text2 + "] " + message;
				File.AppendAllText(_logFilePath, text3 + Environment.NewLine, Encoding.UTF8);
				if (1 == 0)
				{
				}
				ConsoleColor consoleColor = level switch
				{
					LogLevel.TRACE => ConsoleColor.Gray, 
					LogLevel.DEBUG => ConsoleColor.Cyan, 
					LogLevel.INFO => ConsoleColor.White, 
					LogLevel.WARNING => ConsoleColor.Yellow, 
					LogLevel.ERROR => ConsoleColor.Red, 
					LogLevel.FATAL => ConsoleColor.DarkRed, 
					_ => ConsoleColor.White, 
				};
				if (1 == 0)
				{
				}
				ConsoleColor foregroundColor = consoleColor;
				ConsoleColor foregroundColor2 = Console.ForegroundColor;
				Console.ForegroundColor = foregroundColor;
				Console.WriteLine(text3);
				Console.ForegroundColor = foregroundColor2;
			}
		}
		catch (Exception ex)
		{
			Console.WriteLine("❌ Erreur écriture log: " + ex.Message);
		}
	}

	public static void LogException(string context, Exception ex, LogLevel level = LogLevel.ERROR)
	{
		Log("❌ EXCEPTION dans " + context + ":", level);
		Log("   Message: " + ex.Message, level);
		Log("   Type: " + ex.GetType().Name, level);
		if (ex.InnerException != null)
		{
			Log("   Inner Exception: " + ex.InnerException.Message, level);
		}
		Log("   StackTrace:", level);
		string[] array = ex.StackTrace?.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
		if (array != null)
		{
			string[] array2 = array;
			foreach (string text in array2)
			{
				Log("      " + text.Trim(), level);
			}
		}
	}

	public static void Close()
	{
		Log("═══════════════════════════════════════════════════════");
		Log("  SESSION TERMINEE");
		Log("═══════════════════════════════════════════════════════");
	}

	public static void Trace(string message)
	{
		Log(message, LogLevel.TRACE);
	}

	public static void Debug(string message)
	{
		Log(message, LogLevel.DEBUG);
	}

	public static void Info(string message)
	{
		Log(message);
	}

	public static void Warning(string message)
	{
		Log(message, LogLevel.WARNING);
	}

	public static void Error(string message)
	{
		Log(message, LogLevel.ERROR);
	}

	public static void Fatal(string message)
	{
		Log(message, LogLevel.FATAL);
	}
}

