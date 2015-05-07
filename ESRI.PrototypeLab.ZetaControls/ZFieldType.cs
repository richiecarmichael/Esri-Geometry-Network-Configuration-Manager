/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public enum ZFieldType {
        Unknown,
        SmallInteger,
        Integer,
        Single,
        Double,
        String,
        Date,
        OID,
        Geometry,
        Blob,
        Raster,
        GUID,
        GlobalID,
        XML
    }
}
