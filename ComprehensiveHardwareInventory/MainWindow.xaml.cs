using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Linq;
using System.Data;

namespace ComprehensiveHardwareInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region fields
        private string currentfile;
        private static string CurrentProgramPath = System.Environment.CurrentDirectory.Replace("\\bin\\Debug", "\\");
        private string Configfilename = CurrentProgramPath + "Files\\TemplateConfig.xml";
        private string TableToXMLFileName;
        private string persistenFileName;
        private string SystemXPath = "Ecs/ToolControl/Group";
        private string ChamberAXPath = "";
        private string ChamberBXPath = "";
        private DataSet ds = new DataSet();
        #endregion
        public MainWindow()
        {
            InitializeComponent();
            DataTable dt = new DataTable();
            dt.Columns.Add("Channel", typeof(string));
            dt.Columns.Add("Module", typeof(string));
            dt.Columns.Add("Component", typeof(string));
            dt.Columns.Add("Parameter", typeof(string));
            dt.Columns.Add("Anonym", typeof(string));
            dt.Columns.Add("PhysicalAddress", typeof(string));
            dt.Columns.Add("Logic", typeof(string));
            dt.Columns.Add("DateAdded", typeof(string));
            dt.Columns["DateAdded"].DefaultValue = DateTime.Now.ToLocalTime();
            dt.Columns.Add("Tag", typeof(string));
            dt.Columns.Add("Comment", typeof(string));

            ds.Tables.Add(dt);
            ParametersTable.ItemsSource = ds.Tables[0].DefaultView;
            ParametersTable.LoadingRow += new EventHandler<DataGridRowEventArgs>(dataGrid_LoadingRow);
            //OverwriteXMLFile();
            //ReadXML();
            WriteXML();
        }


        private void dataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        private void dataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
           
            if (this.ParametersTable.SelectedItem != null)
            {
                (sender as DataGrid).RowEditEnding -= dataGrid_RowEditEnding;
                (sender as DataGrid).CommitEdit();
                (sender as DataGrid).Items.Refresh();
                (sender as DataGrid).RowEditEnding += dataGrid_RowEditEnding;
            }

            DataRow dgRow = (DataRow)((DataRowView)e.Row.Item).Row;
            //ds.Tables[0].Rows.Add(dgRow);
        }

        #region private methods

        //Common helper methods
        private void GetXMLHierachy()
        {
            
        }

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
        private void OnClickGenerateXML(object sender, RoutedEventArgs e)
        {
            GetXMLHierachy();
            //MoveMotorOperation(1, true);
        }

        private void OnClickSaveToXML(object sender, RoutedEventArgs e)
        {
            GetXMLHierachy();
            //MoveMotorOperation(1, true);
        }
        

        private void OnClickSaveToExcel(object sender, RoutedEventArgs e)
        {
            if (ds.Tables[0].Rows.Count > 0)
            {
                string filename = Tools.SaveExcelFileDialog();
                if (filename.Length > 0)
                {
                    NPOIHelper.ExportDataTableToExcel(ds.Tables[0], filename);
                }
            }      
            else
            {
                MessageBox.Show("table is empty! Try again.");
            }
            
        }
        

        private void OnClickLoadExcel(object sender, RoutedEventArgs e)
        {
            Tuple<string, DataTable> sheets;
            currentfile = Tools.OpenExcelFileDialog();
            if (currentfile != null && currentfile.Length > 0)
            {
                sheets = NPOIHelper.ImportExcelToDataTable(currentfile, true);
                ParametersTable.ItemsSource = sheets.Item2.DefaultView;
                this.Title = currentfile;
            }
            else
            {
                MessageBox.Show("File name empty! Try again.");
            }

        }

        private void OnClickUpdateToExcel(object sender, RoutedEventArgs e)
        {
            DataTable dt = ParametersTable.ItemsSource as DataTable;

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

        // Manipulate XML
        private void WriteXMLNode(string rootNode, double Xvalue, double Zvalue)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(TableToXMLFileName);
            XmlNode root = doc.DocumentElement;
            XmlNode Position = root.SelectSingleNode(rootNode);
            XmlNode XPosition = Position.SelectSingleNode("XPosition");
            XPosition.InnerText = Xvalue.ToString();
            XmlNode ZPosition = Position.SelectSingleNode("ZPosition");
            ZPosition.InnerText = Zvalue.ToString();
            doc.Save(persistenFileName);
        }

        //private void WriteXML(List<string> rowlist)
        private void WriteXML()
        {
            XElement doc = XElement.Load(Configfilename);
            //IEnumerable<XObject> subset = from xobj in doc.Find("System")
            //select xobj;
            XElement System = doc.FindFirstElement("Group","System");

            //XmlDocument doc = XDocument.Load(Configfilename);
            //XElement rootECS = doc.Root.Element("Ecs");
            //XElement sys = rootECS.Element("System");
            //WriteIO(string Index, string Name);
            //WriteComponent(List<string>);
            
            //XmlNode Position = root.SelectSingleNode(SystemXPath);
            //XmlNode XPosition = Position.SelectSingleNode("XPosition");
            //XPosition.InnerText = Xvalue.ToString();
            //XmlNode ZPosition = Position.SelectSingleNode("ZPosition");
            //ZPosition.InnerText = Zvalue.ToString();
            //doc.Save(persistenFileName);
        }

        private XElement MakeAnalogElement(string index,string Name, string Unit, string PhysicalMin, string PhysicalMax, string LogicalMin, string LogicalMax, string LogicOffset)
        {
            string AnalogDirection = index.Substring(0, 2).ToUpper() == "AX" ? "AnaInCell": "AnaOutCell";
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

        private XElement MakeDigitalElement(string index, string Name)
        {
            string DigitalDirection = index.Substring(0, 2).ToUpper() == "DX" ? "DigInCell" : "DigOutCell";
            int indexInt = int.Parse(index.Substring(2));
            XElement result = null;
            if(index.Substring(0, 2).ToUpper() == "DX")
            {
                result = new XElement("DigOutCell",
                            new XElement("Index", indexInt),
                            new XElement("Name", Name),
                            new XElement("Default", "false"),
                            new XElement("NeedLatch", "false"),
                            new XElement("LatchWhen", "false")
                        );
            }
            else if(index.Substring(0, 2).ToUpper() == "DY")
            {
                result = new XElement("DigOutCell",
                            new XElement("Index", indexInt),
                            new XElement("Name", Name)
                        );
            }

            return result;
        }
        private void WriteDataToXML(string path, List<string> dataValue)
        {
            //rowlist[1] rowlist[0]
            //XmlDocument doc = new XmlDocument();
            //doc.Load(TableToXMLFileName);
            //XmlNode root = doc.DocumentElement;
            //XmlNode Position = root.SelectSingleNode(rootNode);
            //XmlNode XPosition = Position.SelectSingleNode("XPosition");
            //XPosition.InnerText = Xvalue.ToString();
            //XmlNode ZPosition = Position.SelectSingleNode("ZPosition");
            //ZPosition.InnerText = Zvalue.ToString();
            //doc.Save(persistenFileName);
        }

        private void OverwriteXMLFile()
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            using (XmlWriter writer = XmlWriter.Create(TableToXMLFileName, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("ToolControl");
                writer.WriteStartElement("Group");
                writer.WriteStartAttribute("Name");
                writer.WriteValue("System");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

        }

        private void ReadXML()
        {
            XElement fromFile = XElement.Load(Configfilename); 
        }

        private void OnClickSaveToConfig(object sender, RoutedEventArgs e)
        {
            foreach (var row in ParametersTable.ItemsSource)
            {
                DataRow r = ((DataRowView)row).Row;
                List<string> list = RowToList(r.ItemArray);
                WriteXML();
            }
            
        }
        

        #endregion
    }

    public class RowObject
    {
        public string IOIndex;
        public string NameFunction;
        public string Anonym;
        public string Logic;
    }

    [Serializable]
    public class RbtPosition
    {
        [XmlElement(ElementName = "XPosition")]
        public double x;
        [XmlElement(ElementName = "ZPosition")]
        public double z;
        public RbtPosition(double X, double Z) { this.x = X; this.z = Z; }
        public RbtPosition() { }
    }

    [Serializable]
    public class RbtPositions
    {
        [XmlElement(ElementName = "LoadPosition")]
        public RbtPosition LoadPosition { get; set; }
        [XmlElement(ElementName = "UnloadPosition")]
        public RbtPosition UnloadPosition { get; set; }
        [XmlElement(ElementName = "SPMPosition")]
        public RbtPosition SPMPosition { get; set; }
        [XmlElement(ElementName = "QDRPosition")]
        public RbtPosition QDRPosition { get; set; }
        public RbtPositions() { }
        public RbtPositions(RbtPosition LoadPos, RbtPosition UnloadPos, RbtPosition SPMPos, RbtPosition QDRPos)
        {
            LoadPosition = LoadPos;
            UnloadPosition = UnloadPos;
            SPMPosition = SPMPos;
            QDRPosition = QDRPos;
        }
    }
}
