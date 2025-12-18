using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Models;

public class FileItem : INotifyPropertyChanged
{
	private bool _isSelected;

	private bool _isInventorFile;

	private string _status = "En attente";

	public string FileName { get; set; }

	public string FileType { get; set; }

	public string FileSizeFormatted { get; set; }

	public string FullPath { get; set; }

	public string FileExtension { get; set; }

	public bool IsSelected
	{
		get
		{
			return _isSelected;
		}
		set
		{
			_isSelected = value;
			OnPropertyChanged("IsSelected");
		}
	}

	public bool IsInventorFile
	{
		get
		{
			return _isInventorFile;
		}
		set
		{
			_isInventorFile = value;
			OnPropertyChanged("IsInventorFile");
		}
	}

	public string Status
	{
		get
		{
			return _status;
		}
		set
		{
			_status = value;
			OnPropertyChanged("Status");
		}
	}

	public event PropertyChangedEventHandler PropertyChanged;

	protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
	{
		this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}
}
