/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZDefaultSubtype : ZSubtype {
        public ZDefaultSubtype() : base() { }
        public ZDefaultSubtype(ZFeatureClass parent) : base(parent, 0, parent.Path.Table) { }
    }
}
