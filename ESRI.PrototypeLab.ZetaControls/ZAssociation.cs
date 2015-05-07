/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZAssociation : ZObject {
        public ZPath Path { get; set; }
        public string Field { get; set; }
    }
}
