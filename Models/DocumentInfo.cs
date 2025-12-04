using System;
using System.ComponentModel;
using QRCodeRevitAddin.Utils;

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
                return ToJson();
            }
        }

        public DocumentInfo()
        {
            _name = string.Empty;
            _revision = string.Empty;
            _project = string.Empty;
            _date = DateValidator.GetTodayFormatted();
            _checkedBy = string.Empty;
        }

        public DocumentInfo(string name, string revision, string project, string date, string checkedBy)
        {
            _name = name ?? string.Empty;
            _revision = revision ?? string.Empty;
            _project = project ?? string.Empty;
            _date = date ?? DateValidator.GetTodayFormatted();
            _checkedBy = checkedBy ?? string.Empty;
        }

        public bool IsValid(out string errorMessage)
        {
            errorMessage = string.Empty;

            if (string.IsNullOrWhiteSpace(Name))
            {
                errorMessage = "Drawing Number is required";
                return false;
            }

            if (string.IsNullOrWhiteSpace(Revision))
            {
                errorMessage = "Sheet Name is required";
                return false;
            }

            if (string.IsNullOrWhiteSpace(Project))
            {
                errorMessage = "Revision is required";
                return false;
            }

            if (string.IsNullOrWhiteSpace(Date))
            {
                errorMessage = "Date is required";
                return false;
            }

            if (!DateValidator.IsValid(Date))
            {
                errorMessage = $"Date must be in format {DateValidator.GetExpectedFormat()}";
                return false;
            }

            if (string.IsNullOrWhiteSpace(CheckedBy))
            {
                errorMessage = "Checked By is required";
                return false;
            }

            return true;
        }

        public string ToJson()
        {
            string escapedName = EscapeJsonString(Name ?? "");
            string escapedRevision = EscapeJsonString(Revision ?? "");
            string escapedProject = EscapeJsonString(Project ?? "");
            string escapedDate = EscapeJsonString(Date ?? "");
            string escapedCheckedBy = EscapeJsonString(CheckedBy ?? "");

            return $"{{\n  \"drawingNo\": \"{escapedName}\",\n  \"sheetName\": \"{escapedRevision}\",\n  \"revision\": \"{escapedProject}\",\n  \"issueDate\": \"{escapedDate}\",\n  \"checkedBy\": \"{escapedCheckedBy}\"\n}}";
        }

        private string EscapeJsonString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
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