/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Framework;
using ESRI.PrototypeLab.ZetaControls;
using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace ESRI.PrototypeLab.Zeta {
    public class ButtonGeometricNetwork : ESRI.ArcGIS.Desktop.AddIns.Button {
        private GeometricNetworkWindow _geometricNetworkWindow = null;
        //
        // CONSTRUCTOR
        //
        public ButtonGeometricNetwork() {
            //AppDomain.CurrentDomain.AssemblyResolve += (s, e) => {
            //    Assembly assembly = null;
            //    for (int i = 0; i <= AppDomain.CurrentDomain.GetAssemblies().Length - 1; i++) {
            //        if (e.Name.Split(",".ToCharArray())[0].ToUpper() ==
            //            AppDomain.CurrentDomain.GetAssemblies()[i].FullName.Split(",".ToCharArray())[0].ToUpper()) {
            //            assembly = AppDomain.CurrentDomain.GetAssemblies()[i];
            //            break;
            //        }
            //    }
            //    return assembly;
            //};
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) => {
                return AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(
                    a => {
                        string x = e.Name.Split(',')[0].ToUpper(CultureInfo.InvariantCulture);
                        string y = a.FullName.Split(',')[0].ToUpper(CultureInfo.InvariantCulture);
                        return x == y;
                    }
                );
            };
        }
        //
        // METHODS
        //
        protected override void OnClick() {
            try {
                if (this._geometricNetworkWindow == null) {
                    Assembly.Load("Microsoft.Windows.Shell");
                    Assembly.Load("RibbonControlsLibrary");
                    Assembly.Load("ESRI.PrototypeLab.ZetaControls");

                    // Get the ArcMap Window Position
                    IWindowPosition windowPosition = (IWindowPosition)ArcMap.Application;

                    // Create a new window
                    double w = 800;
                    double h = 600;
                    double l = windowPosition.Left + (windowPosition.Width / 2) - (w / 2);
                    double t = windowPosition.Top + (windowPosition.Height / 2) - (h / 2);
                    this._geometricNetworkWindow = new GeometricNetworkWindow() {
                        Left = l,
                        Top = t,
                        Width = w,
                        Height = h,
                        Topmost = true
                    };
                    this._geometricNetworkWindow.Closing += (s, e) => {
                        e.Cancel = true;
                        this._geometricNetworkWindow.Hide();
                    };
                    this._geometricNetworkWindow.Show();
                }
                else {
                    if (this._geometricNetworkWindow.IsVisible) {
                        this._geometricNetworkWindow.Hide();
                    }
                    else {
                        this._geometricNetworkWindow.Show();
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }
    }
}
