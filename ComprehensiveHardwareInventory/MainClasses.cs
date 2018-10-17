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
            List<string> result = new List<string> { Channel ?? "", Module ?? "", Parameter ?? "", Anonym ?? "", PhysicalAddress ?? "", Logic ?? "", PhysicalLogic ?? "", DateAdded ?? "", Tag ?? "", Comment ?? "" };
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
