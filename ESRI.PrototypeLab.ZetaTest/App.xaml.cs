/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS;
using ESRI.ArcGIS.esriSystem;
using System.Windows;

namespace ESRI.PrototypeLab.ZetaTest {
    public partial class App : Application {
        private AoInitialize _aoInitialize = null;
        protected override void OnStartup(StartupEventArgs e) {
            base.OnStartup(e);

            // Binding to ArcGIS Desktop
            RuntimeManager.Bind(ProductCode.Desktop);

            // Get ESRI License
            this._aoInitialize = new AoInitializeClass();

            esriLicenseProductCode[] codes = new esriLicenseProductCode[] {
                esriLicenseProductCode.esriLicenseProductCodeStandard,
                esriLicenseProductCode.esriLicenseProductCodeAdvanced
            };
            bool ok = false;
            foreach (esriLicenseProductCode code in codes) {
                if (this._aoInitialize.IsProductCodeAvailable(code) == esriLicenseStatus.esriLicenseAvailable) {
                    if (this._aoInitialize.Initialize(code) == esriLicenseStatus.esriLicenseCheckedOut) {
                        ok = true;
                        break;
                    }
                }
            }
            if (!ok) {
                MessageBox.Show("Unable to checkout Esri license", "Zeta", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        protected override void OnExit(ExitEventArgs e) {
            base.OnExit(e);

            if (this._aoInitialize != null) {
                this._aoInitialize.Shutdown();
            }
        }
    }
}
