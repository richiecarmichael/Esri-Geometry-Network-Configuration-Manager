/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.ObjectModel;

/* 

  A geometric network can have 0..M weights.
  A weight can have 1..M assocations.
  An assocation is a featureclass field.
  A featureclass field can only be associated once in a network.
  A weight cannot have two or more assocations from the same featureclass.
  BitGateSize is always 0 except for weight type of "BitGate"
 
  Weight Type	    Valid Feature Attribute Types
  -----------------------------------------------
  esriWTNull	    None
  esriWTBitgate	    esriFieldTypeSmallInteger, esriFieldTypeInteger
  esriWTInteger	    esriFieldTypeSmallInteger, esriFieldTypeInteger
  esriWTSingle	    esriFieldTypeSingle
  esriWTDouble	    esriFieldTypeSingle, esriFieldTypeDouble
 
 */

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZNetWeight : ZObject {
        //
        // PROPERTIES
        //
        public string Name { get; set; }
        public ZWeightType WeightType { get; set; }
        public int BitGateSize { get; set; }
        public ObservableCollection<ZNetWeightAssocation> NetWeightAssocations { get; private set; }
        //
        // CONSTRUCTOR
        //
        public ZNetWeight() {
            this.NetWeightAssocations = new ObservableCollection<ZNetWeightAssocation>();
        }
        public ZNetWeight(ZGeometricNetwork geometricNetwork, INetWeight netWeight, IEnumNetWeightAssociation enumNetWeightAssociation) {
            // Initialize collections
            this.NetWeightAssocations = new ObservableCollection<ZNetWeightAssocation>();

            // Weight name
            this.Name = netWeight.WeightName;

            // Weight type
            this.WeightType = netWeight.WeightType.ToZWeightType();
            if (this.WeightType == ZWeightType.BitGate) {
                this.BitGateSize = netWeight.BitGateSize;
            }

            // Add network assocations
            INetWeightAssociation netWeightAssocation = enumNetWeightAssociation.Next();
            while (netWeightAssocation != null) {
                this.NetWeightAssocations.Add(new ZNetWeightAssocation(geometricNetwork, netWeightAssocation));
                netWeightAssocation = enumNetWeightAssociation.Next();
            }
        }
    }
}
