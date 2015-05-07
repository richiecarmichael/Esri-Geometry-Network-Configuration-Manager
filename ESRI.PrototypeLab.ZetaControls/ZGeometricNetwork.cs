/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace ESRI.PrototypeLab.ZetaControls {
    [Serializable]
    public class ZGeometricNetwork : ZDataset {
        //
        // CONSTRUCTOR
        //
        public ZGeometricNetwork() {
            // Initialize collections
            this.NetworkClasses = new ObservableCollection<ZNetworkClass>();
            this.Weights = new ObservableCollection<ZNetWeight>();
            this.EdgeRules = new ObservableCollection<ZRule>();
            this.JunctionRules = new ObservableCollection<ZRule>();
        }
        public ZGeometricNetwork(IGeometricNetwork geometricNetwork) : base(geometricNetwork as IDataset)  {
            // Initialize collections
            this.NetworkClasses = new ObservableCollection<ZNetworkClass>();
            this.Weights = new ObservableCollection<ZNetWeight>();
            this.EdgeRules = new ObservableCollection<ZRule>();
            this.JunctionRules = new ObservableCollection<ZRule>();

            // Report addition
            GeometricNetworkViewModel.Default.AddMessage(string.Format("Finding network classes..."), MessageType.Information);

            // Add network classes
            IEnumFeatureClass enumFeatureClass1 = geometricNetwork.get_ClassesByType(esriFeatureType.esriFTSimpleJunction);
            IEnumFeatureClass enumFeatureClass2 = geometricNetwork.get_ClassesByType(esriFeatureType.esriFTSimpleEdge);
            IEnumFeatureClass enumFeatureClass3 = geometricNetwork.get_ClassesByType(esriFeatureType.esriFTComplexEdge);
            foreach (IEnumFeatureClass enumFeatureClass in new IEnumFeatureClass[] { enumFeatureClass1, enumFeatureClass2, enumFeatureClass3 }) {
                IFeatureClass featureClass = enumFeatureClass.Next();
                while (featureClass != null) {
                    INetworkClass networkClass = featureClass as INetworkClass;
                    if (networkClass != null) {
                        // Create network class
                        ZNetworkClass znc = new ZNetworkClass(networkClass);

                        // Add class to collection
                        this.NetworkClasses.Add(znc);

                        // Report addition
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding class: {0}", znc.Path.Table), MessageType.Information);
                    }
                    featureClass = enumFeatureClass.Next();
                }
            }

            // Add weights and weight associations
            INetwork network = geometricNetwork.Network;
            INetSchema netSchema = (INetSchema)network;
            for (int i = 0; i < netSchema.WeightCount; i++) {
                // Create network weight
                INetWeight netWeight = netSchema.get_Weight(i);
                IEnumNetWeightAssociation enumNetWeightAssocation = netSchema.get_WeightAssociations(i);
                ZNetWeight znw = new ZNetWeight(this, netWeight, enumNetWeightAssocation);

                // Add weight to collection
                this.Weights.Add(znw);

                // Report addition
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding weight: {0}", znw.Name), MessageType.Information);
            }

            // Add connectivity rules
            IEnumRule rules = geometricNetwork.Rules;
            IRule rule = rules.Next();
            while (rule != null) {
                if (rule is IJunctionConnectivityRule) {
                    // Create junction rule
                    IJunctionConnectivityRule junctionConnectivityRule = (IJunctionConnectivityRule)rule;
                    ZJunctionConnectivityRule jcr = new ZJunctionConnectivityRule(this, junctionConnectivityRule);

                    // Add junction rule
                    this.JunctionRules.Add(jcr);

                    // Report
                    GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding junction rule: {0}", jcr.Id), MessageType.Information);
                }
                else if (rule is IEdgeConnectivityRule) {
                    // Create edge rule
                    IEdgeConnectivityRule edgeConnectivityRule = (IEdgeConnectivityRule)rule;
                    ZEdgeConnectivityRule ecr = new ZEdgeConnectivityRule(this, edgeConnectivityRule);

                    // Add edge rule
                    this.EdgeRules.Add(ecr);

                    // Report
                    GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding edge rule: {0}", ecr.Id), MessageType.Information);
                }

                rule = rules.Next();
            }
        }
        //
        // PROPERTIES
        //
        public ObservableCollection<ZNetworkClass> NetworkClasses { get; private set; }
        public ObservableCollection<ZNetWeight> Weights { get; private set; }
        public ObservableCollection<ZRule> EdgeRules { get; private set; }
        public ObservableCollection<ZRule> JunctionRules { get; private set; }
        //
        // METHODS
        //
        public ZSubtype FindSubtype(string guid) {
            return this.NetworkClasses.SelectMany(n => n.Subtypes).FirstOrDefault(s => s.Zid == guid);
        }
        public ZSubtype FindSubtype(int featureclass, int subtypecode) {
            var f = this.NetworkClasses.FirstOrDefault(n => n.Id == featureclass);
            if (f == null) { return null; }
            return f.Subtypes.FirstOrDefault(s => s.Code == subtypecode);
        }
    }
}
