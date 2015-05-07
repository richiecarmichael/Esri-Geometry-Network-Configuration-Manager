/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.ComponentModel;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public abstract class ZObject : INotifyPropertyChanged {
        public string Zid { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZObject() : base() {
            this.Zid = Guid.NewGuid().ToString("N").ToUpperInvariant();
        }

        [field: NonSerialized]
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) {
            if (PropertyChanged != null) {
                this.PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
            GeometricNetworkViewModel.Default.MakeDirty();
        }
    }
}
