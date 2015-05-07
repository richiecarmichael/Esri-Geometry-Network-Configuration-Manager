/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZSubtype : ZObject {
        //
        // PROPERTIES
        //
        public int Code { get; private set; }
        public string Name { get; private set; }
        public ZFeatureClass Parent { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZSubtype() : base() { }
        public ZSubtype(ZFeatureClass parent, int code, string name)
            : base() {
            this.Parent = parent;
            this.Code = code;
            this.Name = name;
        }
    }
}
