/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.Windows.Markup;

namespace ESRI.PrototypeLab.ZetaControls {
    [MarkupExtensionReturnType(typeof(object[]))]
    public class EnumValuesExtension : MarkupExtension {
        public EnumValuesExtension() { }
        public EnumValuesExtension(Type enumType) {
            this.EnumType = enumType;
        }

        [ConstructorArgument("enumType")]
        public Type EnumType { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider) {
            if (this.EnumType == null) { return null; }
            return Enum.GetValues(this.EnumType);
        }
    }
}
