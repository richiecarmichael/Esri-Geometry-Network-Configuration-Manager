/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.Linq;
using ESRI.ArcGIS.Geodatabase;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZNetWeightAssocation : ZObject {
        //
        // PROPERTIES
        //
        public ZNetworkClass NetworkClass { get; private set; }
        public ZField Field { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZNetWeightAssocation() { }
        public ZNetWeightAssocation(ZGeometricNetwork geometricNetwork, INetWeightAssociation netWeightAssociation) {
            this.NetworkClass = geometricNetwork.NetworkClasses.FirstOrDefault(n => n.Path.Table == netWeightAssociation.TableName);
            if (this.NetworkClass == null) { return; }
            this.Field = this.NetworkClass.Fields.FirstOrDefault(f => f.Name == netWeightAssociation.FieldName);
        }
    }
}
