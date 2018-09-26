using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Linq;
using System.Reflection;
using System.ComponentModel;

namespace ComprehensiveHardwareInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region fields
        private string currentfile;
        private ListCollectionView view;
        private static string ProjectFilePath = System.Environment.CurrentDirectory.Replace("\\bin\\Debug", "\\") + "Files\\";
        private string Configfilename = ProjectFilePath + "TemplateConfig.xml";
        private string ParameterWordsFile = ProjectFilePath + "ParameterWordsList.txt"; 
        private string LogicWordsFile = ProjectFilePath + "LogicWordsList.txt";
        private string ModuleWordsFile = ProjectFilePath + "ModulesList.txt";
        private ObservableCollection<ItemRow> TableRowsList = new ObservableCollection<ItemRow>();
        private bool IsExcelLoaded = false;
        private string user = ProjectFilePath.Split('\\')[2];
        private string currentIOView = String.Empty;
        //private string currentModuleView = String.Empty;
        private static DataTable dt = new DataTable();
        private static DataTable currentDt = dt.Clone();
        private string currentCellOldValue = String.Empty;
        private List<string> ChannelsListStrings = new List<string> { "AX", "AY", "DX", "DY" };
        private List<string> PhysicalLogicsListStrings = new List<string> { "0-16383", "0-32767"};
        private List<string> ModulesListStrings = new List<string> { "System", "A", "B", "C", "D", "E", "F", "G"};
        private List<string> ParametersListStrings = new List<string> { "CHEM1 Temperature Reading", "FlowReading", "PressureReading", "Yellow", "Green","Blue", "Enable", "UpValve", "OnOffValve", "DownValve", "DryValve", "OutValve", "SupplyValve", "TankToChamberValve", "Valve", "DSP Tank H2O2 In Flow Reading", "VMS Tank Supply Pump Speed Reading",
                                                                        "ReclaimToTankValve", "ExchangerPCWOutValve", "PumpOnOffValve", " FeedbackValve", "FinishAudiableSignal", "StartAudiableSignal", "Signal", "Sensor", "Anneal1 Heater Temperature Reading", "H2SO4 Tank DIW In Flow Reading", "CHEM1 Tank High Sensor", "N2 Protect Bearing Pressure Sensor",  "Wafer Pick Up Position Sensor", "Frame Door Sensor Chamber A1 Backside NO", "Frame Door Sensor Chamber A2 Rightside NO",
                                                                        "H2 MFC Inlet Pressure Reading", "CHEM2", "H2", "CO2", "N2 Line1 MFC Reading", "H2SO4 Supply Levitronix Pump Speed Reading", "DIW", "CDIW Pressure Reading", "DSP", "Heater", "LightTower", "ChamberLight", "FrameLight", "MotorInterlock", "EFEM Interlock And Enable Feedback", "EnvironmentExhaust", "H2O2Mixer",
                                                                        "OuterShroud", "MiddleShroud", "InnerShroud", "Loadport", "MainVacuum", "EFEMIonbarRemotePower", "FacilityCDIW", "", "N2ProtectBearingPCW", "N2PickupPin", "CassetteLot", "Interlock", "MotorInterlock", "Vacuum Pump Interlock And Enable Feedback",
                                                                        "Door", "Pressure", "Leak", "Level", "DSP Cabinet Exhaust Pressure Sensor#1", "DSP Cabinet Leak#1", "Module C Interlock Status", "Module C Door Status", "Heartbeat Interlock Feedback","Process Robot Interlock Interlock And Enable FeedBack"};
        private List<string> LogicsListStrings = new List<string> {  "4-20mA : 0-10LPM", "4-20mA : 0-124.5Pa", "4-20mA : 0-500Pa", "1-5V : 10-100LPM", "4-20mA : 0-0.8Mpa", "4~20MA : 0.2-1.0MΩ·CM", "4~20mA : -15~150PSI",
                                                                     "0~5V : 0-10000RPM", "4~20mA : 0~4.0L/Min ", "4~20mA : 0.0~ -101.3KPa", "open:1", "close:1", "interlock:0", "interlock:1", "leak:0", "leak:1", "on:1", "off:1", "alarm:0", "alarm:1",
                                                                     "level achieved : 0", "level achieved : 1", "overfilled : 0", "overfilled : 1", "normal : 1", "normal : 0", "Up Pos : 1", "Dw Pos : 1", "enalbe:1", "request:1", "ready:1" };

        XElement ToolControl = null;
        XElement doc = null;
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        #endregion

        #region Properties
        Dictionary<string, XElement> HierarchyDic { get; set; } = new Dictionary<string, XElement>();
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            ReadXMLConfig();
            PrepareTables();

            ReadInAutoWordsList(ParameterWordsFile);        // Auto words file watcher process
            ReadInAutoWordsList(LogicWordsFile);
            ReadInAutoWordsList(ModuleWordsFile);
            MyFileSystemWatcher FileWatcher = new MyFileSystemWatcher(ProjectFilePath, "*.txt");
            FileWatcher.OnChanged += new FileSystemEventHandler(OnFileChanged);
            FileWatcher.Start();
        }

        private void PrepareTables()
        {
            dt.Columns.Add("Channel", typeof(string));
            dt.Columns.Add("Module", typeof(string));
            //dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("Parameter", typeof(string));
            dt.Columns.Add("Anonym", typeof(string));
            dt.Columns.Add("PhysicalAddress", typeof(string));
            dt.Columns.Add("Logic", typeof(string));
            dt.Columns.Add("PhysicalLogic", typeof(string));
            dt.Columns.Add("DateAdded", typeof(string));
            dt.Columns.Add("Tag", typeof(string));
            dt.Columns.Add("Comment", typeof(string));


            ParametersTable.ItemsSource = TableRowsList;

            view = (ListCollectionView)CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
            //view.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Descending));
            ChannelTypeGrouper grouper = new ChannelTypeGrouper();
            view.GroupDescriptions.Add(new PropertyGroupDescription("Channel", grouper));

            ParametersTable.LoadingRow += new EventHandler<DataGridRowEventArgs>(dataGrid_LoadingRow);
            ParametersTable.PreviewKeyDown += dataGrid_PreviewKeyDown;
            ParametersTable.RowEditEnding += dataGrid_RowEditEnding;
            ParametersTable.PreparingCellForEdit += dataGrid_PreparingCellForEdit;
            ParametersTable.CurrentCellChanged += dataGrid_CurrentCellChanging;
        }

        private void ReadXMLConfig()
        {
            if (File.Exists(Configfilename))
            {
                doc = XElement.Load(Configfilename);
                ToolControl = doc.FindFirstElement("ToolControl");
                List<string> ModulesName = new List<string>();
                string ModuleName = String.Empty;

                foreach (var module in ToolControl.Elements())
                {
                    ModuleName = module.Attribute("Name").Value;
                    ModulesName.Add(ModuleName);
                    XElement IODefinitionNode = module.FindFirstElement("Property", "AIDefinitions");
                    HierarchyDic.Add((ModuleName + "ax").ToLower(), IODefinitionNode);     // e.g. <systemax,XElement("System/AIDefinitions")>.
                    IODefinitionNode = module.FindFirstElement("Property", "AODefinitions");
                    HierarchyDic.Add((ModuleName + "ay").ToLower(), IODefinitionNode);
                    IODefinitionNode = module.FindFirstElement("Property", "DIDefinitions");
                    HierarchyDic.Add((ModuleName + "dx").ToLower(), IODefinitionNode);
                    IODefinitionNode = module.FindFirstElement("Property", "DODefinitions");
                    HierarchyDic.Add((ModuleName + "dy").ToLower(), IODefinitionNode);
                }
            }
        }
        private void OnFileChanged(object source, FileSystemEventArgs e)      // update list if file has been changed.
        {
            ReadInAutoWordsList(e.FullPath);
        }

        private void dataGrid_CurrentCellChanging(object sender, EventArgs e)
        {
            if (ParametersTable.SelectedItem != null && ParametersTable.SelectedItem != CollectionView.NewItemPlaceholder)
            {
                int i = ParametersTable.SelectedIndex;
                if (i > -1)
                {
                    DataGridRow dataGridRow = ParametersTable.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;
                    if (dataGridRow != null)
                        (ParametersTable.Columns[7].GetCellContent(dataGridRow) as TextBlock).Text = user + " " + DateTime.Now.ToString();
                }
            }
        }

        private void dataGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            DataGridTemplateColumn col = e.Column as DataGridTemplateColumn;
            if (col != null)
            {
                ContentPresenter contentPresenter = e.EditingElement as ContentPresenter;
                DataTemplate editingTemplate = contentPresenter.ContentTemplate;
                AutoCompleteBox ComponentsAutoBox = editingTemplate.FindName("ComponentAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                //if(ComponentsAutoBox != null)
                //{
                //    ComponentsAutoBox.ItemsSource = ComponentsListStrings;
                //    Keyboard.Focus(ComponentsAutoBox);
                //    ComponentsAutoBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                //} 
                //else
                //{
                AutoCompleteBox ParametersAutoBox = editingTemplate.FindName("ParameterAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                if (ParametersAutoBox != null)
                {
                    ParametersAutoBox.ItemsSource = ParametersListStrings;
                    Keyboard.Focus(ParametersAutoBox);
                    ParametersAutoBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                }
                else
                {
                    AutoCompleteBox LogicsAutoBox = editingTemplate.FindName("LogicAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                    if (LogicsAutoBox != null)
                        LogicsAutoBox.ItemsSource = LogicsListStrings;
                    else
                    {
                        AutoCompleteBox ChannelsAutoBox = editingTemplate.FindName("ChannelAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                        if (ChannelsAutoBox != null)
                            ChannelsAutoBox.ItemsSource = ChannelsListStrings;
                        else
                        {
                            AutoCompleteBox ModulesAutoBox = editingTemplate.FindName("ModuleAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                            if (ModulesAutoBox != null)
                                ModulesAutoBox.ItemsSource = ModulesListStrings;
                            else
                            {
                                AutoCompleteBox PhysicalLogicsAutoBox = editingTemplate.FindName("PhysicalLogicAutoCompleteBox", contentPresenter) as AutoCompleteBox;
                                if (PhysicalLogicsAutoBox != null)
                                    PhysicalLogicsAutoBox.ItemsSource = PhysicalLogicsListStrings;
                            }
                        }

                    }
                }
                //} 
            }
        }

        private void ReadInAutoWordsList(string filename)
        {
            string line;
            StreamReader sr = new StreamReader(filename);
            switch (filename.Split('\\').Last())
            {
                case "ParameterWordsList.txt":
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrEmpty(line)){
                                string s = textInfo.ToTitleCase(line.Trim());
                                if (!ParametersListStrings.Contains(s))
                                    ParametersListStrings.Add(s);
                            }

                        }
                    }
                    break;
                case "LogicWordsList.txt":
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrEmpty(line)){
                                string s = textInfo.ToTitleCase(line.Trim());
                                if (!LogicsListStrings.Contains(s))
                                    LogicsListStrings.Add(s);
                            }

                        }
                    }
                    break;
                case "ModulesList.txt":
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!String.IsNullOrEmpty(line))
                            {
                                string s = textInfo.ToTitleCase(line.Trim());
                                if (!ModulesListStrings.Contains(s))
                                    ModulesListStrings.Add(s);
                            }

                        }
                    }
                    break;
            }
        }

        private static object GetCellValue(DataGridCellInfo cell)
        {
            var boundItem = cell.Item;
            var binding = new Binding();
            if (cell.Column is DataGridTextColumn)
            {
                binding
                  = ((DataGridTextColumn)cell.Column).Binding
                        as Binding;
            }
            else if (cell.Column is DataGridCheckBoxColumn)
            {
                binding
                  = ((DataGridCheckBoxColumn)cell.Column).Binding
                        as Binding;
            }
            else if (cell.Column is DataGridComboBoxColumn)
            {
                binding
                    = ((DataGridComboBoxColumn)cell.Column).SelectedValueBinding
                         as Binding;

                if (binding == null)
                {
                    binding
                      = ((DataGridComboBoxColumn)cell.Column).SelectedItemBinding
                           as Binding;
                }
            }

            if (binding != null)
            {
                var propertyName = binding.Path.Path;
                var propInfo = boundItem.GetType().GetProperty(propertyName);
                return propInfo.GetValue(boundItem, new object[] { });
            }

            return null;
        }

        private void dataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {

            if (ParametersTable.SelectedItem != null)
            {
                (sender as DataGrid).RowEditEnding -= dataGrid_RowEditEnding;
                (sender as DataGrid).CommitEdit();
                if (ParametersTable.SelectedItem == CollectionView.NewItemPlaceholder)
                {
                    int i = ParametersTable.SelectedIndex == 0 ? 1 : ParametersTable.SelectedIndex;
                    ItemRow row = ParametersTable.Items.GetItemAt(i - 1) as ItemRow;
                    if (row != null && row.Parameter != null && row.Parameter.IndexOf("Shroud") >= 0)
                    {
                        row.Tag = "X";
                        ParametersTable.SelectedItem = row;
                        (sender as DataGrid).CommitEdit();
                    }
                }
                (sender as DataGrid).Items.Refresh();
                (sender as DataGrid).RowEditEnding += dataGrid_RowEditEnding;
            }

        }

        #region private methods

        //Common helper methods

        private List<string> RowToList(object[] array)
        {
            List<string> result = new List<string>();
            foreach (var item in array)
            {
                string a = item.ToString();    // if the cell is never edited, i.e., null, then convert to ""  automactically
                result.Add(a);
            }
            return result;
        }


        private void OnClickSaveToNewExcel(object sender, RoutedEventArgs e)
        {
            ConvertToDataTable(TableRowsList);
            if (IsExcelLoaded)
            {
                string filename = Tools.SaveExcelFileDialog();
                if (filename.Length > 0)
                {
                    if (NPOIHelper.ExportDataTableToExcel(dt, filename).Item1)
                        MessageBox.Show("successfully save to excel file");
                    else
                        MessageBox.Show("Save failed! check again");

                }
            }
            else if (TableRowsList.Count > 0)
            {
                string filename = Tools.SaveExcelFileDialog();
                if (filename.Length > 0)
                {
                    NPOIHelper.ExportDataTableToExcel(dt, filename);
                    MessageBox.Show("successfully save to excel file");
                }
            }
            else
            {
                MessageBox.Show("table is empty! Try again.");
            }

        }

        private void OnClickModuleType(object sender, RoutedEventArgs e)
        {
            ConvertToDataTable(TableRowsList);
            TableRowsList = ConvertToStringList(dt);
            ParametersTable.ItemsSource = null;
            ParametersTable.ItemsSource = TableRowsList;
            view = (ListCollectionView)CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
            view.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Ascending));
            ModuleTypeGrouper grouper = new ModuleTypeGrouper();
            view.GroupDescriptions.Add(new PropertyGroupDescription("Module", grouper));

            //currentModuleView = (sender as MenuItem).Header.ToString();
            //ParametersTable.ItemsSource = null;
            //ParametersTable.ItemsSource = TableRowsList;
            //ListCollectionView view1 = (ListCollectionView)CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
            ////view1.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Ascending));
            //view1.Filter = new Predicate<object>(item => ((ItemRow)item).Module.ToUpper().Contains(currentIOView));
        }

        private void OnClickIOType(object sender, RoutedEventArgs e)
        {
            currentIOView = (sender as MenuItem).Header.ToString();
            ParametersTable.ItemsSource = null;
            ParametersTable.ItemsSource = TableRowsList;
            ListCollectionView view1 = (ListCollectionView)CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
            //view1.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Ascending));
            view1.Filter = new Predicate<object>(item => ((ItemRow)item).Channel.ToUpper().Contains(currentIOView));
        }
        
        private void OnClickDeleteRow(object sender, RoutedEventArgs e)
        {
            if (ParametersTable.SelectedItem != null && ParametersTable.SelectedItem != CollectionView.NewItemPlaceholder)
            {
                int i = ParametersTable.SelectedIndex;
                ItemRow itemRow = (ItemRow)ParametersTable.SelectedItem;
                if (itemRow != null && itemRow.Channel != null)
                    TableRowsList.Remove(itemRow);
            }
        }

        private void OnClickNormalView(object sender, RoutedEventArgs e)
        {
            ConvertToDataTable(TableRowsList);
            TableRowsList = ConvertToStringList(dt);
            ParametersTable.ItemsSource = null;
            ParametersTable.ItemsSource = TableRowsList;
            view = (ListCollectionView)CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
            //view.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Ascending));
            ChannelTypeGrouper grouper = new ChannelTypeGrouper();
            view.GroupDescriptions.Add(new PropertyGroupDescription("Channel", grouper));
        }

        private void OnClickLoadExcel(object sender, RoutedEventArgs e)
        {
            currentfile = Tools.OpenExcelFileDialog();
            if (currentfile != null && currentfile.Length > 0)
            {
                Tuple<string, DataTable> sheets = NPOIHelper.ImportSheetsToDataTable(currentfile, true);
                dt = NPOIHelper.ImportExcelToDataTable(currentfile, true).Item2;
                //TableRowsList.Clear();
                TableRowsList = ConvertToStringList(dt);
                ParametersTable.ItemsSource = null;
                ParametersTable.ItemsSource = TableRowsList;
                ICollectionView view = CollectionViewSource.GetDefaultView(ParametersTable.ItemsSource);
                //view.SortDescriptions.Add(new SortDescription("Channel", ListSortDirection.Ascending));
                ChannelTypeGrouper grouper = new ChannelTypeGrouper();
                view.GroupDescriptions.Add(new PropertyGroupDescription("Channel", grouper));

                this.Title = currentfile;
                IsExcelLoaded = true;
            }
            else
            {
                MessageBox.Show("File name empty! Try again.");
            }

        }

        private void OnClickUpdateToExcel(object sender, RoutedEventArgs e)
        {
            if (IsExcelLoaded)
            {
                dt.Clear();
                ConvertToDataTable(TableRowsList);

                Tuple<bool, string> result = NPOIHelper.ExportDataTableToExcel(dt, currentfile);
                if (result.Item1)
                {
                    MessageBox.Show("Successfully update to excel");
                }
                else
                {
                    MessageBox.Show("Fail to update to excel! Check again.");
                }
            }
            else
            {
                MessageBox.Show("table is not loaded! Try again.");
            }
        }

        // Manipulate XML
        private void WriteXML(List<string> rowlist)
        {
            XElement doc = XElement.Load(Configfilename);
            XElement System = doc.FindFirstElement("Group", "System");
            XElement ChamA1 = doc.FindFirstElement("Group", "A1");
            XElement ChamA2 = doc.FindFirstElement("Group", "A2");
            XElement ChamB1 = doc.FindFirstElement("Group", "B1");
            XElement ChamB2 = doc.FindFirstElement("Group", "B2");
            XElement ModuleNode = null;
            XElement NewIONode = null;
            XElement IOListNode = null;
            switch (rowlist[1])
            {
                case "System":
                    ModuleNode = System;
                    break;
                case "A1":
                    ModuleNode = ChamA1;
                    break;
                case "A2":
                    ModuleNode = ChamA2;
                    break;
                case "B1":
                    ModuleNode = ChamB1;
                    break;
                case "B2":
                    ModuleNode = ChamB2;
                    break;
            }
            string IOType = rowlist[0].Substring(0, 2).ToUpper();
            switch (IOType)
            {
                case "AX":
                    {
                        NewIONode = MakeAnalogElement(rowlist);
                        IOListNode = ModuleNode.FindFirstElement("Property", "AIDefinitions");
                    }
                    break;
                case "AY":
                    {
                        NewIONode = MakeAnalogElement(rowlist);
                        IOListNode = ModuleNode.FindFirstElement("Property", "AODefinitions");
                    }
                    break;
                case "DX":
                    {
                        NewIONode = MakeDigitalElement(rowlist[0], rowlist[3]);
                        IOListNode = ModuleNode.FindFirstElement("Property", "DIDefinitions");
                    }
                    break;
                case "DY":
                    {
                        NewIONode = MakeDigitalElement(rowlist[0], rowlist[3]);
                        IOListNode = ModuleNode.FindFirstElement("Property", "DODefinitions");
                    }
                    break;
            }

            if (IOListNode != null)
            {
                bool IsFound = false;
                foreach (var item in IOListNode.Elements())
                {
                    if (item.Element("Index").Value == NewIONode.Element("Index").Value && item.Element("Name").Value != NewIONode.Element("Name").Value)
                    {
                        IsFound = true;
                        item.ReplaceAll(NewIONode.Elements());
                        break;
                    }
                }
                if (!IsFound)
                    IOListNode.Add(NewIONode);
            }
            doc.Save(Configfilename);
        }

        private XElement MakeAnalogElement(string index, string Name, string Unit, string PhysicalMin, string PhysicalMax, string LogicalMin, string LogicalMax, string LogicOffset)
        {
            string AnalogDirection = index.Substring(0, 2).ToUpper() == "AX" ? "AnaInCell" : "AnaOutCell";
            int indexInt = int.Parse(index.Substring(2));
            XElement result = new XElement(AnalogDirection,
                                        new XElement("Index", indexInt),
                                        new XElement("Name", Name),
                                        new XElement("Unit", Unit),
                                        new XElement("PhysicalMin", PhysicalMin),
                                        new XElement("PhysicalMax", PhysicalMax),
                                        new XElement("LogicalMin", LogicalMin),
                                        new XElement("LogicalMax", LogicalMax),
                                        new XElement("LogicOffset", LogicOffset)
                                        );
            return result;
        }

        private XElement MakeAnalogElement(List<string> list)
        {
            string Unit;
            string PhysicalMin;
            string PhysicalMax;
            string LogicalMin = String.Empty;
            string LogicalMax = String.Empty;
            string LogicOffset = "0";
            string Name = list[2];

            if (list[5].Trim() == "/10")
            {
                Unit = "C";
                LogicalMax = "100";
                LogicalMin = "0";
                PhysicalMax = "1000";
                PhysicalMin = "0";
            }
            else
            {
                string[] RangeItems = list[5].Split(':');           //e.g. list[5] is 4-20mA:0-32767:0-4.0L/Min
                string[] physicalItems = RangeItems[1].Trim().Split('-');
                PhysicalMax = physicalItems[1].Trim();
                PhysicalMin = physicalItems[0].Trim();
                string[] logicalItems = RangeItems[2].Trim().Split('-');
                LogicalMin = logicalItems[0].Trim();
                LogicalMax = Regex.Replace(logicalItems[1], "[a-z]", "", RegexOptions.IgnoreCase);
                Unit = logicalItems[1].Replace(LogicalMax, "");
            }
            string AnalogDirection = list[0].Substring(0, 2).ToUpper() == "AX" ? "AnaInCell" : "AnaOutCell";
            int indexInt = int.Parse(list[0].Substring(2));
            XElement result = new XElement(AnalogDirection,
                                        new XElement("Index", indexInt),
                                        new XElement("Name", Name),
                                        new XElement("Unit", Unit),
                                        new XElement("PhysicalMin", PhysicalMin),
                                        new XElement("PhysicalMax", PhysicalMax),
                                        new XElement("LogicalMin", LogicalMin),
                                        new XElement("LogicalMax", LogicalMax),
                                        new XElement("LogicOffset", LogicOffset)
                                        );
            return result;
        }

        private XElement MakeDigitalElement(string index, string Name)
        {
            string DigitalDirection = index.Substring(0, 2).ToUpper() == "DX" ? "DigInCell" : "DigOutCell";
            int indexInt = int.Parse(index.Substring(2));
            XElement result = null;
            if (index.Substring(0, 2).ToUpper() == "DX")
            {
                result = new XElement("DigOutCell",
                            new XElement("Index", indexInt),
                            new XElement("Name", Name),
                            new XElement("Default", "false"),
                            new XElement("NeedLatch", "false"),
                            new XElement("LatchWhen", "false")
                        );
            }
            else if (index.Substring(0, 2).ToUpper() == "DY")
            {
                result = new XElement("DigOutCell",
                            new XElement("Index", indexInt),
                            new XElement("Name", Name)
                        );
            }

            return result;
        }


        private void OnClickUpdateConfig(object sender, RoutedEventArgs e)     // This will only update the list, not delete all, change a node if it has already there. 
        {
            //XElement ModuleNode = null;
            XElement NewIONode = null;
            XElement IOListNode = null;

            foreach (var row in ParametersTable.ItemsSource)
            {
                ItemRow r = (ItemRow)row;
                List<string> list = r.ToList();

                //ModuleNode = ToolControl.Element(textInfo.ToTitleCase(list[1]));

                if (r.Channel != null && r.Channel.Length > 2)
                {
                    string IOType = list[0].Substring(0, 2).ToUpper();
                    switch (IOType)
                    {
                        case "DX":
                        case "DY":
                            NewIONode = MakeDigitalElement(list[0], list[3]);
                            break;
                        case "AX":
                        case "AY":
                            NewIONode = MakeAnalogElement(list);
                            break;
                        default:
                            MessageBox.Show("IO index {0} is wrong! check again.", list[0]);
                            return;
                    }
                    IOListNode = HierarchyDic[(list[1].Trim() + list[0].Trim().Substring(0, 2)).ToLower()];   // e.g. "System" + "Ax"  => systemax .

                    if (IOListNode != null)
                    {
                        foreach (var item in IOListNode.Elements())
                        {
                            if (item.Element("Index").Value == NewIONode.Element("Index").Value)
                            {
                                if(item.Element("Name").Value != NewIONode.Element("Name").Value)
                                    item.ReplaceAll(NewIONode.Elements());
                                break;
                            }
                            else if (int.Parse(item.Element("Index").Value) > int.Parse(NewIONode.Element("Index").Value))
                            {
                                item.AddBeforeSelf(NewIONode);
                                break;
                            }
                        }

                    }
                }
            }

            doc.Save(Configfilename);
            MessageBox.Show("Successfully update to xml config!");
        }

        private void OnClickOverwriteConfig(object sender, RoutedEventArgs e)   // This will delete all the IO list and then add from the beginning.
        {
            foreach (var item in HierarchyDic.Values)              // Here is the differences with the OnClickSaveToConfig method, only this loop.
            {
                item.RemoveAll();
            }
            //XElement ModuleNode = null;
            XElement NewIONode = null;
            XElement IOListNode = null;

            foreach (var row in ParametersTable.ItemsSource)
            {
                DataRow r = ((DataRowView)row).Row;
                List<string> list = RowToList(r.ItemArray);

                //ModuleNode = ToolControl.Element(textInfo.ToTitleCase(list[1]));

                string IOType = list[0].Substring(0, 2).ToUpper();
                switch (IOType)
                {
                    case "DX":
                    case "DY":
                        NewIONode = MakeDigitalElement(list[0], list[3]);
                        break;
                    case "AX":
                    case "AY":
                        NewIONode = MakeAnalogElement(list);
                        break;
                    default:
                        MessageBox.Show("IO index {0} is wrong! check again.", list[0]);
                        return;
                }
                IOListNode = HierarchyDic[(list[1].Trim() + list[0].Trim().Substring(0, 2)).ToLower()];   // e.g. "System" + "Ax"  => systemax .

                if (IOListNode != null)
                {
                    bool IsFound = false;
                    foreach (var item in IOListNode.Elements())
                    {
                        if (item.Element("Index").Value == NewIONode.Element("Index").Value && item.Element("Name").Value != NewIONode.Element("Name").Value)
                        {
                            IsFound = true;
                            item.ReplaceAll(NewIONode.Elements());
                            break;
                        }
                    }
                    if (!IsFound)
                        IOListNode.Add(NewIONode);
                }
            }

            doc.Save(Configfilename);
            MessageBox.Show("Successfully update to xml config!");
        }

        private static DataGridCell GetCell(DataGrid dataGrid, DataGridRow rowContainer, int column)
        {
            if (rowContainer != null)
            {
                DataGridCellsPresenter presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
                if (presenter != null)
                    return presenter.ItemContainerGenerator.ContainerFromIndex(column) as DataGridCell;
            }

            return null;
        }

        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }

        private void dataGrid_PreviewKeyDown(object sender, KeyEventArgs e)  // This will set focus on first column of the new unedited row
        {                                                                    // This is very hard to achieved since the cell in the new row wouldn't be created yet.
            DataGrid grid = (DataGrid)sender;

            if (e.Key == Key.Enter || e.Key == Key.Return)
            {
                DataGridRow row = ParametersTable.ItemContainerGenerator.ContainerFromItem(CollectionView.NewItemPlaceholder) as DataGridRow;
                if (row != null)
                {
                    if (row.GetIndex() == ParametersTable.SelectedIndex + 1)
                    {
                        ParametersTable.SelectedItem = row.DataContext;
                        DataGridCell cell = GetCell(ParametersTable, row, 0);
                        if (cell != null)
                        {
                            ParametersTable.CurrentCell = new DataGridCellInfo(cell);
                            ParametersTable.BeginEdit();       // make the newly empty place holder into edit mode.

                        }
                    }
                    else
                    {
                        DataGridCell cell = GetCell(ParametersTable, ParametersTable.SelectedItem as DataGridRow, 0);
                        if (cell != null)
                        {
                            ParametersTable.CurrentCell = new DataGridCellInfo(cell);
                            ParametersTable.BeginEdit();       // make the newly empty place holder into edit mode.

                        }
                    }
                }

            }
        }
        #endregion

        private void OnClickNew(object sender, RoutedEventArgs e)
        {
            IsExcelLoaded = false;
            dt.Rows.Clear();
            TableRowsList.Clear();
            this.Title = "No file is loaded";
        }

        public static ObservableCollection<ItemRow> ConvertToStringList(DataTable table)
        {
            if (table == null)
            {
                return null;
            }
            ObservableCollection<ItemRow> result = new ObservableCollection<ItemRow>();
            string s = String.Empty;
            foreach (DataRow row in table.Rows)
            {
                ItemRow itemrow = new ItemRow(row.ItemArray[0].ToString(), row.ItemArray[1].ToString(), row.ItemArray[2].ToString(), row.ItemArray[3].ToString(), row.ItemArray[4].ToString(), row.ItemArray[5].ToString(), row.ItemArray[6].ToString(), row.ItemArray[7].ToString(), row.ItemArray[8].ToString(), row.ItemArray[9].ToString());

                result.Add(itemrow);
            }

            return result;
        }

        public static DataTable ConvertToDataTable(ObservableCollection<ItemRow> rowlist)
        {
            if (rowlist.Count == 0)
            {
                return null;
            }
            string s = String.Empty;
            dt.Clear();                           // Or we can create another datatable here and then merge with dt.
            foreach (var row in rowlist.ToArray())
            {
                if (!String.IsNullOrEmpty(row.Channel))
                {
                    DataRow dtrow = dt.NewRow();
                    dtrow[0] = row.Channel;
                    dtrow[1] = row.Module;
                    //dtrow[2] = row.Component;
                    dtrow[2] = row.Parameter;
                    dtrow[3] = row.Anonym;
                    dtrow[4] = row.PhysicalAddress;
                    dtrow[5] = row.Logic;
                    dtrow[6] = row.PhysicalLogic;
                    dtrow[7] = row.DateAdded;
                    dtrow[8] = row.Tag;
                    dtrow[9] = row.Comment;
                    dt.Rows.Add(dtrow);
                }
            }

            return dt;

        }

        private void OnClickAddAutoWords(object sender, RoutedEventArgs e)
        {
            //FileStream fs = new FileStream(AutoWordsFile, FileMode.Append);
            //string filename = Tools.SaveExcelFileDialog();
            string header = (sender as MenuItem).Header.ToString().Substring(3).TrimEnd('W','o','r','d','s');
            string file = ParameterWordsFile;     // open ParameterWordsFile by default
            switch (header)
            {
                case "Logic":
                    file = LogicWordsFile;
                    break;
                case "Module":
                    file = ModuleWordsFile;
                    break;
            }
           System.Diagnostics.Process.Start("notepad.exe", file);
        }
    }

    public class ItemRow
    {
        public string Channel { get; set; }
        public string Module { get; set; }
        //public string Component { get; set; }
        public string Parameter { get; set; }
        
        public string Anonym { get; set; }
        public string PhysicalAddress { get; set; }
        public string Logic { get; set; }
        public string PhysicalLogic { get; set; }
        public string DateAdded { get; set; }
        public string Tag { get; set; }
        public string Comment { get; set; }
        public ItemRow(string channel, string module, string parameter, string anonym, string physicaladdress, string logic, string physicallogic, string dateadded, string tag, string comment)
        {
            Channel = channel;
            Module = module;
            //Component = component;
            Parameter = parameter;
            Anonym = anonym;
            PhysicalAddress = physicaladdress;
            Logic = logic;
            PhysicalLogic = physicallogic;
            DateAdded = dateadded;
            Tag = tag;
            Comment = comment;
        }
        public ItemRow() { }
        public List<string> ToList()
        {
            List<string> result = new List<string> { Channel ?? "", Module ?? "", Parameter ?? "", Anonym ?? "", PhysicalAddress ?? "", Logic ?? "", PhysicalLogic ?? "", DateAdded ?? "", Tag ?? "", Channel ?? "" };
            return result;
        }
    }

    public class ChannelTypeGrouper : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }
            else
            {
                string s = value.ToString();
                if (!String.IsNullOrEmpty(s))
                {
                    if (s.Length > 2)
                    {
                        string s2 = s.Substring(0, 2).ToUpper();
                        switch (s2)
                        {
                            case "AX":
                            case "AY":
                            case "DX":
                            case "DY":
                                return String.Format(culture, s2);
                        }
                    }
                }
                return "Error";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("This converter is for grouping only");
        }
    }

    public class ModuleTypeGrouper : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }
            else
            {
                string s = value.ToString();
                if (!String.IsNullOrEmpty(s))
                {
                   return String.Format(culture, s);
                }
                return "Error";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("This converter is for grouping only");
        }
    }

    public class AutoCompleteFocusableBox : AutoCompleteBox
    {
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            var textbox = Template.FindName("Text", this) as TextBox;
            if (textbox != null) textbox.Focus();
        }

        public new void Focus()
        {
            var textbox = Template.FindName("Text", this) as TextBox;
            if (textbox != null) textbox.Focus();
        }
    }

}
