/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using ESRI.ArcGIS.Geodatabase;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZField : ZObject {
        //
        // PROPERTIES
        //
        public ZFeatureClass Parent { get; private set; }
        public string Name { get; private set; }
        public string Alias { get; private set; }
        public ZFieldType Type { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZField() : base() { }
        public ZField(ZFeatureClass parent, IField field)
            : base() {

            this.Parent = parent;
            this.Name = field.Name;
            this.Alias = field.AliasName;
            this.Type = field.Type.ToZFieldType();
        }
    }
}
