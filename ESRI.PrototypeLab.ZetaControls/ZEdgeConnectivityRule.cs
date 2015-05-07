/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.ObjectModel;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZEdgeConnectivityRule : ZRule {
        private ZSubtype _defaultJunction = null;
        //
        // CONSTRUCTOR
        //
        public ZEdgeConnectivityRule()
            : base() {
            // Initialize collection
            this.Junctions = new ObservableCollection<ZSubtype>();
        }
        public ZEdgeConnectivityRule(ZGeometricNetwork geometricNetwork, IEdgeConnectivityRule edgeConnectivityRule)
            : base(edgeConnectivityRule as IRule) {
            // Initialize collection
            this.Junctions = new ObservableCollection<ZSubtype>();

            // Get From Edge
            ZSubtype zSubtypeEdge1 = geometricNetwork.FindSubtype(edgeConnectivityRule.FromEdgeClassID, edgeConnectivityRule.FromEdgeSubtypeCode);
            this.FromEdge = zSubtypeEdge1;

            // Get To Edge
            ZSubtype zSubtypeEdge2 = geometricNetwork.FindSubtype(edgeConnectivityRule.ToEdgeClassID, edgeConnectivityRule.ToEdgeSubtypeCode);
            this.ToEdge = zSubtypeEdge2;

            // Loop for each junction
            for (int i = 0; i < edgeConnectivityRule.JunctionCount; i++) {
                // Get junction feature class
                ZSubtype zSubtypeJunction = geometricNetwork.FindSubtype(edgeConnectivityRule.get_JunctionClassID(i), edgeConnectivityRule.get_JunctionSubtypeCode(i));

                // Add junction
                this.Junctions.Add(zSubtypeJunction);
            }

            // Store default
            this.DefaultJunction = geometricNetwork.FindSubtype(edgeConnectivityRule.DefaultJunctionClassID, edgeConnectivityRule.DefaultJunctionSubtypeCode);
        }
        public ZEdgeConnectivityRule(ZSubtype fromEdge, ZSubtype toEdge)
            : this() {
            this.FromEdge = fromEdge;
            this.ToEdge = toEdge;
        }
        //
        // PROPERTIES
        //
        public ZSubtype FromEdge { get; private set; }
        public ZSubtype ToEdge { get; private set; }
        public ObservableCollection<ZSubtype> Junctions { get; private set; }
        public ZSubtype DefaultJunction {
            get {
                return this._defaultJunction;
            }
            set {
                this._defaultJunction = value;
                this.OnPropertyChanged("DefaultJunction");
            }
        }
    }
}