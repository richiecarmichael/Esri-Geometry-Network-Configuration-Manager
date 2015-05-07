/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;

namespace ESRI.PrototypeLab.ZetaControls {
    public static class RuleMatrix {
        //
        // CONSTANT
        //
        public const string EDGE_FEATURECLASS = "EdgeFeatureClass";
        public const string EDGE_SUBTYPE = "EdgeSubtype";
        //
        // STATIC METHODS
        //
        public static IList ToJunctionRuleMatrix(ZGeometricNetwork gn) {
            // Create temporary assembly
            AssemblyName an = new AssemblyName("TempAssembly" + gn.GetHashCode());
            AssemblyBuilder assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            TypeBuilder tb = moduleBuilder.DefineType(
                "TempType" + gn.GetHashCode(),
                TypeAttributes.Public |
                TypeAttributes.Class |
                TypeAttributes.AutoClass |
                TypeAttributes.AnsiClass |
                TypeAttributes.BeforeFieldInit |
                TypeAttributes.AutoLayout,
                typeof(object)
            );

            // Add edge fields (first column)
            RuleMatrix.CreateProperty(tb, RuleMatrix.EDGE_FEATURECLASS, typeof(ZNetworkClass));

            // Add edge fields (second column)
            RuleMatrix.CreateProperty(tb, RuleMatrix.EDGE_SUBTYPE, typeof(ZSubtype));

            // Add junction fields (columns)
            foreach (ZNetworkClass nc in gn.NetworkClasses.Where(n => n.IsJunction).OrderBy(n => n.Path.Table)) {
                foreach(ZSubtype st in nc.Subtypes.OrderBy(s => s.Name)){
                    RuleMatrix.CreateProperty(tb, st.Zid, typeof(ZRule));
                }
            }

            // Add edge subtypes (rows)
            Type objectType = tb.CreateType();
            Type listType = typeof(List<>).MakeGenericType(new[] { objectType });
            IList list = Activator.CreateInstance(listType) as IList;
            foreach (ZNetworkClass nc in gn.NetworkClasses.Where(n => n.IsEdge).OrderBy(n => n.Path.Table)) {
                foreach (ZSubtype st in nc.Subtypes.OrderBy(s => s.Name)) {
                    // Create row
                    var row = Activator.CreateInstance(objectType);

                    // Edge feature class
                    PropertyInfo property = objectType.GetProperty(RuleMatrix.EDGE_FEATURECLASS);
                    property.SetValue(
                        row,
                        nc,
                        null);

                    // Edge subtype
                    PropertyInfo property2 = objectType.GetProperty(RuleMatrix.EDGE_SUBTYPE);
                    property2.SetValue(
                        row,
                        st,
                        null);

                    // Add row to list
                    list.Add(row);
                }
            }

            // Add junction rule data
            foreach (var row in list) {
                ZSubtype subtype = objectType.GetProperty(RuleMatrix.EDGE_SUBTYPE).GetValue(row, null) as ZSubtype;
                var rules = gn.JunctionRules
                    .Select(r => { return r as ZJunctionConnectivityRule; })
                    .Where(r => r.Edge.Zid == subtype.Zid);

                foreach (ZJunctionConnectivityRule rule in rules) {
                    PropertyInfo property2 = objectType.GetProperty(rule.Junction.Zid);
                    property2.SetValue(
                        row,
                        rule,
                        null);
                };
            }

            return list;
        }
        public static IList ToEdgeRuleMatrix(ZGeometricNetwork gn) {
            // Create temporary assembly
            AssemblyName an = new AssemblyName("TempAssembly" + gn.GetHashCode());
            AssemblyBuilder assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            TypeBuilder tb = moduleBuilder.DefineType(
                "TempType" + gn.GetHashCode(),
                TypeAttributes.Public |
                TypeAttributes.Class |
                TypeAttributes.AutoClass |
                TypeAttributes.AnsiClass |
                TypeAttributes.BeforeFieldInit |
                TypeAttributes.AutoLayout,
                typeof(object)
            );

            // Add edge featureclass (first column)
            RuleMatrix.CreateProperty(tb, RuleMatrix.EDGE_FEATURECLASS, typeof(ZNetworkClass));

            // Add edge subtype (second column)
            RuleMatrix.CreateProperty(tb, RuleMatrix.EDGE_SUBTYPE, typeof(ZSubtype));

            // Add duplicate featureclass/subtype fields (columns)
            foreach (ZNetworkClass nc in gn.NetworkClasses.Where(n => n.IsEdge).OrderBy(n => n.Path.Table)) {
                foreach (ZSubtype st in nc.Subtypes.OrderBy(s => s.Name)) {
                    RuleMatrix.CreateProperty(tb, st.Zid, typeof(ZRule));
                }
            }

            // Add edge subtypes (rows)
            Type objectType = tb.CreateType();
            Type listType = typeof(List<>).MakeGenericType(new[] { objectType });
            IList list = Activator.CreateInstance(listType) as IList;
            foreach (ZNetworkClass nc in gn.NetworkClasses.Where(n => n.IsEdge).OrderBy(n => n.Path.Table)) {
                foreach (ZSubtype st in nc.Subtypes.OrderBy(s => s.Name)) {
                    // Create row
                    var row = Activator.CreateInstance(objectType);
                    
                    // Edge feature class
                    PropertyInfo property = objectType.GetProperty(RuleMatrix.EDGE_FEATURECLASS);
                    property.SetValue(
                        row,
                        nc,
                        null);

                    // Edge subtype
                    PropertyInfo property2 = objectType.GetProperty(RuleMatrix.EDGE_SUBTYPE);
                    property2.SetValue(
                        row,
                        st,
                        null);

                    // Add row to list
                    list.Add(row);
                }
            }

            //
            foreach (var row in list) {
                ZSubtype subtype = objectType.GetProperty(RuleMatrix.EDGE_SUBTYPE).GetValue(row, null) as ZSubtype;
                var rules = gn.EdgeRules
                    .Select(r => { return r as ZEdgeConnectivityRule; })
                    .Where(r => r.FromEdge.Zid == subtype.Zid);

                foreach (ZEdgeConnectivityRule rule in rules) {
                    PropertyInfo property2 = objectType.GetProperty(rule.ToEdge.Zid);
                    property2.SetValue(
                        row,
                        rule,
                        null);
                };
            }

            return list;
        }
        private static void CreateProperty(TypeBuilder tb, string propertyName, Type propertyType) {
            FieldBuilder fieldBuilder = tb.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);
            PropertyBuilder propertyBuilder = tb.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);
            MethodBuilder getPropMthdBldr =
                tb.DefineMethod(
                    "get_" + propertyName,
                    MethodAttributes.Public |
                    MethodAttributes.SpecialName |
                    MethodAttributes.HideBySig,
                    propertyType, Type.EmptyTypes
                );

            ILGenerator getIL = getPropMthdBldr.GetILGenerator();
            getIL.Emit(OpCodes.Ldarg_0);
            getIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getIL.Emit(OpCodes.Ret);

            MethodBuilder setPropMthdBldr =
                tb.DefineMethod(
                    "set_" + propertyName,
                  MethodAttributes.Public |
                  MethodAttributes.SpecialName |
                  MethodAttributes.HideBySig,
                  null, new Type[] { propertyType });

            ILGenerator setIL = setPropMthdBldr.GetILGenerator();
            setIL.Emit(OpCodes.Ldarg_0);
            setIL.Emit(OpCodes.Ldarg_1);
            setIL.Emit(OpCodes.Stfld, fieldBuilder);
            setIL.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getPropMthdBldr);
            propertyBuilder.SetSetMethod(setPropMthdBldr);
        }
    }
}
