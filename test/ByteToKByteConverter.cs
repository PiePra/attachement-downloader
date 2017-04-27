using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace test
{
    class ByteToKByteConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (Double.TryParse(value.ToString(), out double val))
            {
                if (val < 1000) return (double)value + "B";
                else if (val < 1000000) return Format((double)value / 1000) + "KB";
                else if (val < 1000000000) return Format((double)value / 1000000) + "MB";
                else if (val < 1000000000000) return Format((double)value / 1000000000) + "GB";
                else if (val < 1000000000000000) return Format((double)value / 1000000000000) + "TB";
                else if (val < 1000000000000000000) return Format((double)value / 1000000000000000) + "PB";
                else return Format((double)value / 1000000000000000000) + "ZB";
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string s = value.ToString();
            double val = -1;
            switch (s.Substring(s.Length - 2))
            {
                case "KB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000;
                    break;
                case "MB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000000;
                    break;
                case "GB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000000000;
                    break;
                case "TB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000000000000;
                    break;
                case "PB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000000000000000;
                    break;
                case "ZB":
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    val *= 1000000000000000000;
                    break;
                default: //Byte
                    Double.TryParse(s.Substring(0, s.Length - 3), out val);
                    break;
            }
            return val;
        }

        private string Format(double value)
        {
            var s = string.Format("{0:0.00}", value);
            if (s.EndsWith("00")) return ((int)value).ToString();
            else return s;
        }

    }
}
