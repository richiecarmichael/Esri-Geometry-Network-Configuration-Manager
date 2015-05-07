/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public abstract class ZDataset : ZObject {
        public ZDataset() { }
        public ZDataset(IDataset dataset) {
            this.Path = new ZPath(dataset);
        }
        public ZPath Path { get; set; }
    }
}
