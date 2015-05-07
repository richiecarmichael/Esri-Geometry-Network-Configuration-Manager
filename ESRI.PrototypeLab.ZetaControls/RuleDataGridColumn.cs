/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ESRI.PrototypeLab.ZetaControls {
    public class RuleDataGridColumn : DataGridTemplateColumn {
        public string ColumnName { get; set; }
        protected override FrameworkElement GenerateElement(DataGridCell cell, object dataItem) {
            FrameworkElement frameworkElement = base.GenerateElement(cell, dataItem);
            Binding binding = new Binding(this.ColumnName) {
                Source = dataItem
            };
            frameworkElement.SetBinding(ContentControl.ContentProperty, binding);
            return frameworkElement;
        }
        protected override FrameworkElement GenerateEditingElement(DataGridCell cell, object dataItem) {
            FrameworkElement frameworkElement = base.GenerateEditingElement(cell, dataItem);
            Binding binding = new Binding(this.ColumnName) {
                Source = dataItem
            };
            frameworkElement.SetBinding(ContentControl.ContentProperty, binding);
            return frameworkElement;
        }
    }
}
