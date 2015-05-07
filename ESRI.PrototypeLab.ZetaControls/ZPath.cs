/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using ESRI.ArcGIS.Geodatabase;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZPath : ZObject {
        //
        // PROPERTIES
        //
        public string Database { get; private set; }
        public string Owner { get; private set; }
        public string Table { get; set; }
        //
        // CONSTRUCTOR
        //
        public ZPath(IDataset dataset) {
            IWorkspace workspace = dataset.Workspace;
            ISQLSyntax sqlSyntax = (ISQLSyntax)workspace;
            string database;
            string owner;
            string table;
            sqlSyntax.ParseTableName(dataset.Name, out database, out owner, out table);
            this.Database = database;
            this.Owner = owner;
            this.Table = table;
        }
    }
}
