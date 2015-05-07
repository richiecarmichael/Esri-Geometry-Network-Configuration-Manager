/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using System;

namespace ESRI.PrototypeLab.ZetaControls {
    public class GeometricNetworkLoader  {
        private readonly IName _name = null;
        //
        // CONSTRUCTOR
        //
        public GeometricNetworkLoader(byte[] bytes) {
            // Cast byte to object
            object obj = (object)bytes;

            // Unpack dropped object to Esri name enumerator
            INameFactory nameFactory = new NameFactoryClass();
            IEnumName enumName = nameFactory.UnpackageNames(ref obj);
            this._name = enumName.Next();
        }
        public GeometricNetworkLoader(IName name) {
            this._name = name;
        }
        //
        // METHODS
        //
        public void Load() {
            try {
                // Check parsed named object
                if (this._name == null) {
                    throw new Exception("Invalid object");
                }

                // Open geometric network
                IGeometricNetwork geometricNework = null;
                try {
                    geometricNework = this._name.Open() as IGeometricNetwork;
                }
                catch {
                    throw new Exception("Cannot open geometric network");
                }
                if (geometricNework == null) {
                    throw new Exception("Not a geometric network");
                }

                // Create geometric network
                ZGeometricNetwork zgn = new ZGeometricNetwork(geometricNework);
                GeometricNetworkViewModel.Default.Dataset = zgn;
                GeometricNetworkViewModel.Default.CreateRuleDataSource();
                GeometricNetworkViewModel.Default.AddMessage("Load completed successfully", MessageType.Information);
            }
            catch (Exception ex) {
                // Construct error message
                string message = ex.Message;
#if DEBUG
                message += ex.StackTrace;
#endif
                // Add error message
                if (!string.IsNullOrEmpty(message)) {
                    GeometricNetworkViewModel.Default.AddMessage(message, MessageType.Error);
                }
            }
        }
    }
}