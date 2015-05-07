/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System.Windows;
using System.Windows.Controls;
using ESRI.ArcGIS.Geodatabase;

namespace ESRI.PrototypeLab.ZetaControls {
    public static class Extensions {
        public static void ScrollIntoView(this ItemsControl control, object item) {
            FrameworkElement framework = control.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement;
            if (framework == null) { return; }
            framework.BringIntoView();
        }
        public static void ScrollIntoView(this ItemsControl control) {
            int count = control.Items.Count;
            if (count == 0) { return; }
            object item = control.Items[count - 1];
            control.ScrollIntoView(item);
        }
        public static IGeometricNetwork FindGeometricNetwork(this IWorkspace workspace, string name) {
            IEnumDataset datasets = workspace.get_Datasets(esriDatasetType.esriDTFeatureDataset);
            IDataset dataset = datasets.Next();
            while (dataset != null) {
                INetworkCollection nc = (INetworkCollection)dataset;
                IGeometricNetwork gn = null;
                try {
                    gn = nc.get_GeometricNetworkByName(name);
                }
                catch { }
                if (gn != null) {
                    return gn;
                }
                dataset = datasets.Next();
            }
            return null;
        }
        public static IFeatureClass FindFeatureClass(this IWorkspace workspace, string name) {
            IFeatureWorkspace fw = (IFeatureWorkspace)workspace;
            IFeatureClass fc1 = null;
            try {
                fc1 = fw.OpenFeatureClass(name);
            }
            catch { }
            if (fc1 != null) {
                return fc1;
            }

            IEnumDataset datasets2 = workspace.get_Datasets(esriDatasetType.esriDTFeatureDataset);
            IDataset dataset2 = datasets2.Next();
            while (dataset2 != null) {
                IFeatureClassContainer fcc = (IFeatureClassContainer)dataset2;
                IFeatureClass fc = null;
                try {
                    fc = fcc.get_ClassByName(name);
                }
                catch { }
                if (fc != null) {
                    return fc;
                }
                dataset2 = datasets2.Next();
            }

            return null;
        }
        public static ZFieldType ToZFieldType(this esriFieldType fieldType) {
            switch (fieldType) {
                case esriFieldType.esriFieldTypeBlob:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeDate:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeDouble:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeGeometry:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeGlobalID:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeGUID:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeInteger:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeOID:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeRaster:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeSingle:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeSmallInteger:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeString:
                    return ZFieldType.Blob;
                case esriFieldType.esriFieldTypeXML:
                    return ZFieldType.Blob;
                default:
                    return ZFieldType.Unknown;
            }
        }
        public static ZWeightType ToZWeightType(this esriWeightType type) {
            switch (type) {
                case esriWeightType.esriWTBitGate:
                    return ZWeightType.BitGate;
                case esriWeightType.esriWTBoolean:
                    return ZWeightType.Boolean;
                case esriWeightType.esriWTDouble:
                    return ZWeightType.Double;
                case esriWeightType.esriWTInteger:
                    return ZWeightType.Integer;
                case esriWeightType.esriWTNull:
                    return ZWeightType.Null;
                case esriWeightType.esriWTSingle:
                    return ZWeightType.Single;
                default:
                    return ZWeightType.Unknown;
            }
        }
        public static esriWeightType ToWeightType(this ZWeightType type) {
            switch (type) {
                case ZWeightType.BitGate:
                    return esriWeightType.esriWTBitGate;
                case ZWeightType.Boolean:
                    return esriWeightType.esriWTBoolean;
                case ZWeightType.Double:
                    return esriWeightType.esriWTDouble;
                case ZWeightType.Integer:
                    return esriWeightType.esriWTInteger;
                case ZWeightType.Null:
                    return esriWeightType.esriWTNull;
                case ZWeightType.Single:
                    return esriWeightType.esriWTSingle;
                default:
                    return esriWeightType.esriWTNull;
            }
        }
        public static int? ToVerfiedId(this ZFeatureClass fc, IWorkspace workspace) {
            IFeatureClass featureclass = workspace.FindFeatureClass(fc.Path.Table);
            if (featureclass == null) {
                return null;
            }
            return featureclass.FeatureClassID;
        }
    }
}
