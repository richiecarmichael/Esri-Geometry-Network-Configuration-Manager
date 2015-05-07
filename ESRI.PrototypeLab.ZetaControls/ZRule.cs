/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZRule : ZObject {
        public int Id { get; set; }
        //
        // CONSTRUCTOR
        //
        public ZRule() : base() { }
        public ZRule(IRule rule) : base() {
            this.Id = rule.ID;
        }
    }
}
