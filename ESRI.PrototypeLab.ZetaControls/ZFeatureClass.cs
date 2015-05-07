/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.ObjectModel;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZFeatureClass : ZDataset {
        //
        // PROPERTIES
        //
        public int Id { get; private set; }
        public ObservableCollection<ZSubtype> Subtypes { get; private set; }
        public ObservableCollection<ZField> Fields { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZFeatureClass() {
            this.Subtypes = new ObservableCollection<ZSubtype>();
            this.Fields = new ObservableCollection<ZField>();
        }
        public ZFeatureClass(IFeatureClass featureClass) : base(featureClass as IDataset) {
            this.Subtypes = new ObservableCollection<ZSubtype>();
            this.Fields = new ObservableCollection<ZField>();
            this.Id = featureClass.FeatureClassID;

            ISubtypes subtypes = (ISubtypes)featureClass;
            if (subtypes.HasSubtype) {
                IEnumSubtype enumSubtype = subtypes.Subtypes;
                int code;
                string subtypeName = enumSubtype.Next(out code);
                while (!string.IsNullOrEmpty(subtypeName)) {
                    this.Subtypes.Add(new ZSubtype(this, code, subtypeName));
                    subtypeName = enumSubtype.Next(out code);
                }
            }
            else {
                this.Subtypes.Add(new ZDefaultSubtype(this));
            }

            IFields fields = featureClass.Fields;
            for (int i = 0; i < fields.FieldCount; i++) {
                IField field = fields.get_Field(i);
                this.Fields.Add(new ZField(this, field));
            }
        }
    }
}
