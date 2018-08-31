using System;
using System.IO;
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
using System.Data;

namespace ComprehensiveHardwareInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region fields
        private List<RowObject> ItemList;
        private string persistenFileName = @"Files/TemplateConfig.Config";
        private string TableToXMLFileName = @"TableConfig.xml";
        private DataSet ds = new DataSet();
        private ExcelHelper excelhelper;
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

            //DataRow row = dt.NewRow();
            //row["Channel"] = "AX0";
            //row["Module"] = "System";
            //row["Component"] = "CHEM1";
            //row["Parameter"] = "TemperatureReading";
            //row["PhysicalAddress"] = "CH #1";
            //row["Logic"] = "/10";
            ////row["DateAdded"] = DateTime.Now.ToLocalTime();
            //row["Tag"] = "N/A";
            //row["Comment"] = "N/A";
            //dt.Rows.Add(row);
            ds.Tables.Add(dt);
            ParametersTable.ItemsSource = ds.Tables[0].DefaultView;
            ParametersTable.LoadingRow += new EventHandler<DataGridRowEventArgs>(dataGrid_LoadingRow);
            OverwriteXMLFile();
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

        }

        #region private methods

        //Common helper methods
        private void GetXMLHierachy()
        {

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
            string filename = String.Empty;
            excelhelper = new ExcelHelper(filename);

            GetXMLHierachy();
            //MoveMotorOperation(1, true);
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

        private void WriteRowToXML(DataRow dr)
        {
            //WriteXMLNode();
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
