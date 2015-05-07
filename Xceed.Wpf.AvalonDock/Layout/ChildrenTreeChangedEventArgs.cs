﻿/************************************************************************

   AvalonDock

   Copyright (C) 2007-2013 Xceed Software Inc.

   This program is provided to you under the terms of the New BSD
   License (BSD) as published at http://avalondock.codeplex.com/license 

   For more features, controls, and fast professional support,
   pick up AvalonDock in Extended WPF Toolkit Plus at http://xceed.com/wpf_toolkit

   Stay informed: follow @datagrid on Twitter or Like facebook.com/datagrids

  **********************************************************************/

using System;

namespace Xceed.Wpf.AvalonDock.Layout {
    public enum ChildrenTreeChange {
        /// <summary>
        /// Direct insert/remove operation has been perfomed to the group
        /// </summary>
        DirectChildrenChanged,

        /// <summary>
        /// An element below in the hierarchy as been added/removed
        /// </summary>
        TreeChanged
    }

    public class ChildrenTreeChangedEventArgs : EventArgs {
        public ChildrenTreeChangedEventArgs(ChildrenTreeChange change) {
            Change = change;
        }

        public ChildrenTreeChange Change { get; private set; }
    }
}
