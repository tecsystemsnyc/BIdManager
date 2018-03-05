using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using Microsoft.Win32;
using System;
using System.Windows.Input;
using Microsoft.Office.Interop;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ClosedXML.Excel;
using System.Windows.Forms;
using System.Linq;
using System.IO;

namespace BidManager.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        private string _bidNumber;

        public string BidNumber
        {
            get { return _bidNumber; }
            set
            {
                _bidNumber = value;
                RaisePropertyChanged("BidNumber");
            }
        }
        
        private string _projectName;

        public string ProjectName
        {
            get { return _projectName; }
            set
            {
                _projectName = value;
                RaisePropertyChanged("ProjectName");
            }
        }
        
        private string _location;

        public string Location
        {
            get { return _location; }
            set
            {
                _location = value;
                RaisePropertyChanged("Location");
            }
        }

        private string _salesperson;

        public string Salesperson
        {
            get { return _salesperson; }
            set
            {
                _salesperson = value;
                RaisePropertyChanged("Salesperson");
            }
        }

        private string _estimator;

        public string Estimator
        {
            get { return _estimator; }
            set
            {
                _estimator = value;
                RaisePropertyChanged("Estimator");
            }
        }

        private DateTime _receivedDate;

        public DateTime ReceivedDate
        {
            get { return _receivedDate; }
            set
            {
                _receivedDate = value;
                RaisePropertyChanged("ReceivedDate");
            }
        }

        private DateTime _dueDate;

        public DateTime DueDate
        {
            get { return _dueDate; }
            set
            {
                _dueDate = value;
                RaisePropertyChanged("DueDate");
            }
        }

        private string _requestedBy;

        public string RequestedBy
        {
            get { return _requestedBy; }
            set
            {
                _requestedBy = value;
                RaisePropertyChanged("RequestedBy");
            }
        }

        private string _client;

        public string Client
        {
            get { return _client; }
            set
            {
                _client = value;
                RaisePropertyChanged("Client");
            }
        }

        private bool _requiresScope;

        public bool RequiresScope
        {
            get { return _requiresScope; }
            set
            {
                _requiresScope = value;
                RaisePropertyChanged("RequiresScope");
            }
        }

        private bool _technicalRequired;

        public bool TechnicalRequired
        {
            get { return _technicalRequired; }
            set
            {
                _technicalRequired = value;
                RaisePropertyChanged("TechnicalRequired");
            }
        }

        private string _bidFolder;

        public string BidFolder
        {
            get { return _bidFolder; }
            set
            {
                _bidFolder = value;
                RaisePropertyChanged("BidFolder");
            }
        }

        private string _bidLog;

        public string BidLog
        {
            get { return _bidLog; }
            set
            {
                _bidLog = value;
                RaisePropertyChanged("BidLog");
            }
        }

        private string _referenceFolder;

        public string ReferenceFolder
        {
            get { return _referenceFolder; }
            set
            {
                _referenceFolder = value;
                RaisePropertyChanged("ReferenceFolder");
            }
        }

        public ICommand SelectBidFolderCommand { get; private set; }
        public ICommand SelectBidLogCommand { get; private set; }
        public ICommand SelectReferenceFolderCommand { get; private set; }
        public ICommand SetupBidCommand { get; private set; }

        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            SelectBidFolderCommand = new RelayCommand(selectBidFolderExecute, selectBidFolderCanExecute);
            SelectBidLogCommand = new RelayCommand(selectBidLogExecute, selectBidLogCanExecute);
            SelectReferenceFolderCommand = new RelayCommand(selectReferenceFolderExecute, selectReferenceFolderCanExecute);
            SetupBidCommand = new RelayCommand(setupBidExecute, canSetupBid);

            ReceivedDate = DateTime.Now;
            DueDate = DateTime.Now;
            ReferenceFolder = Properties.Settings.Default.ReferenceFolder;
            BidLog = Properties.Settings.Default.BidLog;
        }

        private void selectReferenceFolderExecute()
        {
            FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            string path = null;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                path = dialog.SelectedPath;
                Properties.Settings.Default.ReferenceFolder = path;
                Properties.Settings.Default.Save();
            }
            ReferenceFolder = path;
        }

        private bool selectReferenceFolderCanExecute()
        {
            return true;
        }

        private void setupBidExecute()
        {
            setupBid();
        }

        private bool canSetupBid()
        {
            bool hasBidFolder = (BidFolder != null && BidFolder != "");
            bool hasLogFolder = (BidLog != null && BidLog != "");
            bool hasReference = (ReferenceFolder != null && ReferenceFolder != "");
            return (hasBidFolder && hasLogFolder && hasReference);
        }

        private void selectBidLogExecute()
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.DefaultExt = ".xlsx";
            string path = null;
            if (dialog.ShowDialog() == true)
            {
                path = dialog.FileName;
                Properties.Settings.Default.BidLog = path;
                Properties.Settings.Default.Save();
            }
            BidLog = path;
        }

        private bool selectBidLogCanExecute()
        {
            return true;
        }

        private void selectBidFolderExecute()
        {
            Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.FileName = BidNumber + " " + ProjectName;
            string path = null;
            if (dialog.ShowDialog() == true)
            {
                path = dialog.FileName;
            }
            BidFolder = path;
        }

        private bool selectBidFolderCanExecute()
        {
            return true;
        }

        private void setupBid()
        {
            if (IsFileLocked(BidLog))
            {
                MessageBox.Show("Please close the bid log before proceeding.");
            }
            else if(!File.Exists(BidFolder + @"\Proposal Opening Form.docx"))
            {
                MessageBox.Show("Please provide a reference fodler with a propopsal opening form.");
            }
            else
            {
                CopyDir.Copy(ReferenceFolder, BidFolder);
                writeToProposalOpening();
                writeToBidLog(BidLog);
                setupEvent(BidNumber, ProjectName, DueDate);
            }
            
        }
        private void setupEvent(string bidNumber, string name, DateTime due)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.AppointmentItem appointment =
                (Microsoft.Office.Interop.Outlook.AppointmentItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
            appointment.Subject = bidNumber + " - " + name;
            appointment.Start = due;
            appointment.End = due;
            appointment.Save();
        }
        private void writeToProposalOpening()
        {
            Dictionary<string, string> fields = new Dictionary<string, string>();
            fields["<Bid Number>"] = BidNumber;
            fields["<Project Name>"] = ProjectName;
            fields["<Location>"] = Location;
            fields["<Salesperson>"] = Salesperson;
            fields["<Estimator>"] = Estimator;
            fields["<DateReceived>"] = ReceivedDate.ToShortDateString();
            fields["<DateDue>"] = DueDate.ToShortDateString();
            fields["<RequestedBy>"] = RequestedBy;


            using (WordprocessingDocument doc =
                  WordprocessingDocument.Open(BidFolder + @"\Proposal Opening Form.docx", true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                Table table = body.Elements<Table>().First();
                foreach(TableRow row in table.Elements<TableRow>())
                {
                    foreach(TableCell cell in row.Elements<TableCell>())
                    {
                        foreach (var para in cell.Elements<Paragraph>())
                        {
                            foreach (var run in para.Elements<Run>())
                            {
                                foreach (var text in run.Elements<Text>())
                                {
                                    foreach (string key in fields.Keys)
                                    {
                                        if (text.Text.Contains(key))
                                        {
                                            text.Text = text.Text.Replace(key, fields[key]);
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
                
            }
        }
        private void writeToBidLog(string path)
        {
            if (!IsFileLocked(path))
            {
                var wb = new XLWorkbook(path);
                var ws = wb.Worksheet("Bid Log");

                var nextRow = ws.Row(10);
                while (nextRow.Cell(1).Value.ToString() != "" &&
                    (string)nextRow.Cell(2).Value.ToString() != "")
                {
                    nextRow = nextRow.RowBelow();
                }
                nextRow.Cell(1).Value = BidNumber;
                nextRow.Cell(2).Value = ProjectName;
                nextRow.Cell(3).Value = Location;
                nextRow.Cell(4).Value = Client;
                nextRow.Cell(5).Value = Salesperson;
                nextRow.Cell(6).Value = Estimator;
                nextRow.Cell(7).Value = DueDate.ToShortDateString();
                nextRow.Cell(11).Value = RequiresScope ? "Yes" : "No";
                nextRow.Cell(12).Value = TechnicalRequired ? "Yes" : "No";

                wb.Save();
            } else
            {
                MessageBox.Show("Bid log is open elsewhere. Please close and try again.");
            }
           
        }

        public bool IsFileLocked(string filePath)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (!File.Exists(filePath))
            {
                return false;
            }

            FileInfo file = new FileInfo(filePath);

            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}