/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZJunctionConnectivityRule : ZRule {
        private bool _isDefault = false;
        private int _edgeMinimum = -1;
        private int _edgeMaximum = -1;
        private int _junctionMinimum = -1;
        private int _junctionMaximum = -1;
        //
        // CONSTRUCTOR
        //
        public ZJunctionConnectivityRule() : base() { }
        public ZJunctionConnectivityRule(ZSubtype edge, ZSubtype junction)
            : this() {
            this.Edge = edge;
            this.Junction = junction;
        }
        public ZJunctionConnectivityRule(ZGeometricNetwork geometricNetwork, IJunctionConnectivityRule junctionConnectivityRule)
            : base(junctionConnectivityRule as IRule) {
            // Get Edge
            ZSubtype zSubtypeEdge = geometricNetwork.FindSubtype(junctionConnectivityRule.EdgeClassID, junctionConnectivityRule.EdgeSubtypeCode);
            this.Edge = zSubtypeEdge;

            // Get Junction
            ZSubtype zSubtypeJunction = geometricNetwork.FindSubtype(junctionConnectivityRule.JunctionClassID, junctionConnectivityRule.JunctionSubtypeCode);
            this.Junction = zSubtypeJunction;

            // Get Cardinality
            this.EdgeMinimum = junctionConnectivityRule.EdgeMinimumCardinality;
            this.EdgeMaximum = junctionConnectivityRule.EdgeMaximumCardinality;
            this.JunctionMinimum = junctionConnectivityRule.JunctionMinimumCardinality;
            this.JunctionMaximum = junctionConnectivityRule.JunctionMaximumCardinality;

            // Get 
            IJunctionConnectivityRule2 junctionConnectivityRule2 = (IJunctionConnectivityRule2)junctionConnectivityRule;
            this.IsDefault = junctionConnectivityRule2.DefaultJunction;
        }
        //
        // PROPERTIES
        //
        public ZSubtype Edge { get; private set; }
        public ZSubtype Junction { get; private set; }
        public bool IsDefault {
            get { return this._isDefault; }
            set {
                this._isDefault = value;
                this.OnPropertyChanged("IsDefault");
            }
        }
        public int EdgeMinimum {
            get { return this._edgeMinimum; }
            set {
                this._edgeMinimum = value;
                this.OnPropertyChanged("EdgeMinimum");
            }
        }
        public int EdgeMaximum {
            get { return this._edgeMaximum; }
            set {
                this._edgeMaximum = value;
                this.OnPropertyChanged("EdgeMaximum");
            }
        }
        public int JunctionMinimum {
            get { return this._junctionMinimum; }
            set {
                this._junctionMinimum = value;
                this.OnPropertyChanged("JunctionMinimum");
            }
        }
        public int JunctionMaximum {
            get { return this._junctionMaximum; }
            set {
                this._junctionMaximum = value;
                this.OnPropertyChanged("JunctionMaximum");
            }
        }
    }
}
