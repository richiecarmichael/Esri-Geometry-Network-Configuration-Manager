/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using Microsoft.Windows.Controls.Ribbon;
using System;
using System.Windows;
using System.Windows.Controls.Primitives;

namespace ESRI.PrototypeLab.ZetaControls {
    public class CompactRibbonApplicationMenu : RibbonApplicationMenu {
        private const double DefaultPopupWidth = 180;

        public double PopupWidth {
            get { return (double)GetValue(PopupWidthProperty); }
            set { SetValue(PopupWidthProperty, value); }
        }

        public static readonly DependencyProperty PopupWidthProperty =
            DependencyProperty.Register("PopupWidth", typeof(double),
            typeof(CompactRibbonApplicationMenu), new UIPropertyMetadata(DefaultPopupWidth));

        public override void OnApplyTemplate() {
            base.OnApplyTemplate();
            this.DropDownOpened += new EventHandler(this.SlimRibbonApplicationMenu_DropDownOpened);
        }
        private void SlimRibbonApplicationMenu_DropDownOpened(object sender, System.EventArgs e) {
            DependencyObject popupObj = base.GetTemplateChild("PART_Popup");
            Popup popupPanel = (Popup)popupObj;
            popupPanel.Width = (double)GetValue(PopupWidthProperty);
        }
    }
}
