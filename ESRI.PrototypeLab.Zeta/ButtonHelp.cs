/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System.Diagnostics;

namespace ESRI.PrototypeLab.Zeta {
    public class ButtonHelp : ESRI.ArcGIS.Desktop.AddIns.Button {
        public ButtonHelp() { }
        protected override void OnClick() {
            Process process = new Process() {
                StartInfo = new ProcessStartInfo() {
                    FileName = "http://blogs.esri.com/esri/apl/2013/09/11/configuration-manager/",
                    Verb = "Open",
                    CreateNoWindow = false
                }
            };
            process.Start();
        }
        protected override void OnUpdate() { }
    }
}
