/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.Globalization;
using System.Windows.Data;

namespace ESRI.PrototypeLab.ZetaControls {
    public class RuleCardinalityConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            //
            if (value == null) { return null; }
            if (parameter == null) { return null; }

            int c = (int)value;
            string par = (string)parameter;

            switch (par) {
                case "EdgeMinimum":
                    return c == -1 ? "0" : c.ToString();
                case "EdgeMaximum":
                    return c == -1 ? "α" : c.ToString();
                case "JunctionMinimum":
                    return c == -1 ? "0" : c.ToString();
                case "JunctionMaximum":
                    return c == -1 ? "α" : c.ToString();
            }

            return null;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            throw new NotImplementedException();
        }
    }
}
