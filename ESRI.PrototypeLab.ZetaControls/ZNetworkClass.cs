/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZNetworkClass : ZFeatureClass {
        private bool _isSourceSink = false;
        //
        //
        //
        public ZNetworkClass() { }
        public ZNetworkClass(INetworkClass networkClass): base(networkClass as IFeatureClass) {
            // Is Junction?
            this.IsJunction = networkClass.FeatureType == esriFeatureType.esriFTSimpleJunction;

            // Is Edge?
            this.IsEdge = networkClass.FeatureType == esriFeatureType.esriFTSimpleEdge || networkClass.FeatureType == esriFeatureType.esriFTComplexEdge;
            
            // Is Complex?
            this.IsComplex = networkClass.FeatureType == esriFeatureType.esriFTComplexEdge;

            // Is Source or Sink?
            this.IsSourceSink = networkClass.NetworkAncillaryRole == esriNetworkClassAncillaryRole.esriNCARSourceSink;

            // Is Orphan Junction?
            this.IsOrphanJunctionFeatureClass = networkClass.FeatureClassID == networkClass.GeometricNetwork.OrphanJunctionFeatureClass.FeatureClassID;
        }
        public bool IsOrphanJunctionFeatureClass { get; private set; }
        public bool IsJunction { get; private set; }
        public bool IsEdge { get; private set; }
        public bool IsComplex { get; set; }
        public bool IsSourceSink {
            get { return this._isSourceSink; }
            set {
                this._isSourceSink = value;
                this.OnPropertyChanged("IsSourceSink");
            }
        }
    }
}
