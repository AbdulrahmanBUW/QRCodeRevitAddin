using System;
using System.ComponentModel;

namespace QRCodeRevitAddin.Models
{
    /// <summary>
    /// Model class representing the information to be encoded in a QR code.
    /// Implements INotifyPropertyChanged for data binding support.
    /// </summary>
    public class DocumentInfo : INotifyPropertyChanged
    {
        private string _name;
        private string _revision;
        private string _project;
        private string _date;

        /// <summary>
        /// Name or sheet number field.
        /// For manual entry: Document/Drawing name.
        /// For sheet extraction: Sheet number.
        /// </summary>
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

        /// <summary>
        /// Revision field.
        /// For manual entry: Custom revision text.
        /// For sheet extraction: Sheet name or current revision.
        /// </summary>
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

        /// <summary>
        /// Project field.
        /// For manual entry: Project name.
        /// For sheet extraction: Additional project information.
        /// </summary>
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

        /// <summary>
        /// Date field in dd/MM/yyyy format.
        /// Defaults to today's date.
        /// </summary>
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

        /// <summary>
        /// Combined text that will be encoded in the QR code.
        /// Format: "{Name} | {Revision} | {Project} | {Date}"
        /// </summary>
        public string CombinedText
        {
            get
            {
                return $"{Name ?? ""} | {Revision ?? ""} | {Project ?? ""} | {Date ?? ""}";
            }
        }

        /// <summary>
        /// Default constructor. Initializes Date to today in dd/MM/yyyy format.
        /// </summary>
        public DocumentInfo()
        {
            _name = string.Empty;
            _revision = string.Empty;
            _project = string.Empty;
            _date = DateTime.Now.ToString("dd/MM/yyyy");
        }

        /// <summary>
        /// Constructor with all fields.
        /// </summary>
        /// <param name="name">Name or sheet number</param>
        /// <param name="revision">Revision text</param>
        /// <param name="project">Project name</param>
        /// <param name="date">Date in dd/MM/yyyy format</param>
        public DocumentInfo(string name, string revision, string project, string date)
        {
            _name = name ?? string.Empty;
            _revision = revision ?? string.Empty;
            _project = project ?? string.Empty;
            _date = date ?? DateTime.Now.ToString("dd/MM/yyyy");
        }

        /// <summary>
        /// Validates that all required fields have values.
        /// </summary>
        /// <returns>True if all fields are non-empty, false otherwise</returns>
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(Name) &&
                   !string.IsNullOrWhiteSpace(Revision) &&
                   !string.IsNullOrWhiteSpace(Project) &&
                   !string.IsNullOrWhiteSpace(Date);
        }

        /// <summary>
        /// Creates a copy of this DocumentInfo object.
        /// </summary>
        /// <returns>A new DocumentInfo with the same values</returns>
        public DocumentInfo Clone()
        {
            return new DocumentInfo(Name, Revision, Project, Date);
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
