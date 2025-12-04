using System;
using System.ComponentModel;

namespace QRCodeRevitAddin.Models
{
    public class DocumentInfo : INotifyPropertyChanged
    {
        private string _name;
        private string _revision;
        private string _project;
        private string _date;
        private string _checkedBy;

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                    OnPropertyChanged(nameof(CombinedText));
                }
            }
        }

        public string Revision
        {
            get => _revision;
            set
            {
                if (_revision != value)
                {
                    _revision = value;
                    OnPropertyChanged(nameof(Revision));
                    OnPropertyChanged(nameof(CombinedText));
                }
            }
        }

        public string Project
        {
            get => _project;
            set
            {
                if (_project != value)
                {
                    _project = value;
                    OnPropertyChanged(nameof(Project));
                    OnPropertyChanged(nameof(CombinedText));
                }
            }
        }

        public string Date
        {
            get => _date;
            set
            {
                if (_date != value)
                {
                    _date = value;
                    OnPropertyChanged(nameof(Date));
                    OnPropertyChanged(nameof(CombinedText));
                }
            }
        }

        public string CheckedBy
        {
            get => _checkedBy;
            set
            {
                if (_checkedBy != value)
                {
                    _checkedBy = value;
                    OnPropertyChanged(nameof(CheckedBy));
                    OnPropertyChanged(nameof(CombinedText));
                }
            }
        }

        public string CombinedText
        {
            get
            {
                return $"{Name ?? ""} | {Revision ?? ""} | {Project ?? ""} | {Date ?? ""} | {CheckedBy ?? ""}";
            }
        }

        public DocumentInfo()
        {
            _name = string.Empty;
            _revision = string.Empty;
            _project = string.Empty;
            _date = DateTime.Now.ToString("dd/MM/yyyy");
            _checkedBy = string.Empty;
        }

        public DocumentInfo(string name, string revision, string project, string date, string checkedBy)
        {
            _name = name ?? string.Empty;
            _revision = revision ?? string.Empty;
            _project = project ?? string.Empty;
            _date = date ?? DateTime.Now.ToString("dd/MM/yyyy");
            _checkedBy = checkedBy ?? string.Empty;
        }

        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(Name) &&
                   !string.IsNullOrWhiteSpace(Revision) &&
                   !string.IsNullOrWhiteSpace(Project) &&
                   !string.IsNullOrWhiteSpace(Date) &&
                   !string.IsNullOrWhiteSpace(CheckedBy);
        }

        public DocumentInfo Clone()
        {
            return new DocumentInfo(Name, Revision, Project, Date, CheckedBy);
        }

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}