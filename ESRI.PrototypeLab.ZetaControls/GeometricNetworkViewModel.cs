/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System.Collections;
using System.Windows;
using System.Windows.Data;

namespace ESRI.PrototypeLab.ZetaControls {   
    public class GeometricNetworkViewModel : ViewModel {
        //
        // CONSTRUCTOR
        //
        private GeometricNetworkViewModel() {
            this.WindowTitle = "Geometric Network Configuration Manager";
        }
        //
        // STATIC
        //
        private static GeometricNetworkViewModel defaultInstance = new GeometricNetworkViewModel();
        public static GeometricNetworkViewModel Default {
            get { return defaultInstance; }
        }
        public static readonly DependencyProperty JunctionRuleDataSourceProperty = DependencyProperty.Register(
            "JunctionRuleDataSource",
            typeof(ListCollectionView),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty EdgeRuleDataSourceProperty = DependencyProperty.Register(
            "EdgeRuleDataSource",
            typeof(ListCollectionView),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty SelectedNetworkClassProperty = DependencyProperty.Register(
            "SelectedNetworkClass",
            typeof(ZNetworkClass),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty SelectedWeightProperty = DependencyProperty.Register(
            "SelectedWeight",
            typeof(ZNetWeight),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty SelectedJunctionRuleProperty = DependencyProperty.Register(
            "SelectedJunctionRule",
            typeof(ZJunctionConnectivityRule),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty SelectedEdgeRuleProperty = DependencyProperty.Register(
            "SelectedEdgeRule",
            typeof(ZEdgeConnectivityRule),
            typeof(GeometricNetworkViewModel),
            new PropertyMetadata(null)
        );
        //
        // PROPERTIES
        //
        public ListCollectionView JunctionRuleDataSource {
            get { return (ListCollectionView)this.GetValue(GeometricNetworkViewModel.JunctionRuleDataSourceProperty); }
            set { this.SetValue(GeometricNetworkViewModel.JunctionRuleDataSourceProperty, value); }
        }
        public ListCollectionView EdgeRuleDataSource {
            get { return (ListCollectionView)this.GetValue(GeometricNetworkViewModel.EdgeRuleDataSourceProperty); }
            set { this.SetValue(GeometricNetworkViewModel.EdgeRuleDataSourceProperty, value); }
        }
        public ZNetworkClass SelectedNetworkClass {
            get { return (ZNetworkClass)this.GetValue(GeometricNetworkViewModel.SelectedNetworkClassProperty); }
            set { this.SetValue(GeometricNetworkViewModel.SelectedNetworkClassProperty, value); }
        }
        public ZNetWeight SelectedWeight {
            get { return (ZNetWeight)this.GetValue(GeometricNetworkViewModel.SelectedWeightProperty); }
            set { this.SetValue(GeometricNetworkViewModel.SelectedWeightProperty, value); }
        }
        public ZJunctionConnectivityRule SelectedJunctionRule {
            get { return (ZJunctionConnectivityRule)this.GetValue(GeometricNetworkViewModel.SelectedJunctionRuleProperty); }
            set { this.SetValue(GeometricNetworkViewModel.SelectedJunctionRuleProperty, value); }
        }
        public ZEdgeConnectivityRule SelectedEdgeRule {
            get { return (ZEdgeConnectivityRule)this.GetValue(GeometricNetworkViewModel.SelectedEdgeRuleProperty); }
            set { this.SetValue(GeometricNetworkViewModel.SelectedEdgeRuleProperty, value); }
        }
        //
        // METHODS
        //
        public void CreateRuleDataSource() {
            IList list = RuleMatrix.ToJunctionRuleMatrix(this.Dataset as ZGeometricNetwork);
            ListCollectionView lc = new ListCollectionView(list);
            lc.GroupDescriptions.Add(new PropertyGroupDescription(RuleMatrix.EDGE_FEATURECLASS + ".Path.Table"));
            this.JunctionRuleDataSource = lc;

            IList list2 = RuleMatrix.ToEdgeRuleMatrix(this.Dataset as ZGeometricNetwork);
            ListCollectionView lc2 = new ListCollectionView(list2);
            lc2.GroupDescriptions.Add(new PropertyGroupDescription(RuleMatrix.EDGE_FEATURECLASS + ".Path.Table"));
            this.EdgeRuleDataSource = lc2;
        }
        public override void Clear() {
            base.Clear();
            this.JunctionRuleDataSource = null;
            this.EdgeRuleDataSource = null;
            this.SelectedJunctionRule = null;
        }
        public override void Load(string document) {
            base.Load(document);
            this.CreateRuleDataSource();
        }
    }
}
