﻿/************************************************************************

   AvalonDock

   Copyright (C) 2007-2013 Xceed Software Inc.

   This program is provided to you under the terms of the New BSD
   License (BSD) as published at http://avalondock.codeplex.com/license 

   For more features, controls, and fast professional support,
   pick up AvalonDock in Extended WPF Toolkit Plus at http://xceed.com/wpf_toolkit

   Stay informed: follow @datagrid on Twitter or Like facebook.com/datagrids

  **********************************************************************/

using System.Linq;
using System.Windows;
using System.Windows.Media;
using Xceed.Wpf.AvalonDock.Layout;

namespace Xceed.Wpf.AvalonDock.Controls {
    internal class DocumentPaneGroupDropTarget : DropTarget<LayoutDocumentPaneGroupControl> {
        internal DocumentPaneGroupDropTarget(LayoutDocumentPaneGroupControl paneControl, Rect detectionRect, DropTargetType type)
            : base(paneControl, detectionRect, type) {
            _targetPane = paneControl;
        }

        LayoutDocumentPaneGroupControl _targetPane;

        protected override void Drop(LayoutDocumentFloatingWindow floatingWindow) {
            ILayoutPane targetModel = _targetPane.Model as ILayoutPane;

            switch (Type) {
                case DropTargetType.DocumentPaneGroupDockInside:
                    #region DropTargetType.DocumentPaneGroupDockInside
 {
                        var paneGroupModel = targetModel as LayoutDocumentPaneGroup;
                        var paneModel = paneGroupModel.Children[0] as LayoutDocumentPane;
                        var sourceModel = floatingWindow.RootDocument;

                        paneModel.Children.Insert(0, sourceModel);
                    }
 break;
                    #endregion
            }
            base.Drop(floatingWindow);
        }

        protected override void Drop(LayoutAnchorableFloatingWindow floatingWindow) {
            ILayoutPane targetModel = _targetPane.Model as ILayoutPane;

            switch (Type) {
                case DropTargetType.DocumentPaneGroupDockInside:
                    #region DropTargetType.DocumentPaneGroupDockInside
 {
                        var paneGroupModel = targetModel as LayoutDocumentPaneGroup;
                        var paneModel = paneGroupModel.Children[0] as LayoutDocumentPane;
                        var layoutAnchorablePaneGroup = floatingWindow.RootPanel as LayoutAnchorablePaneGroup;

                        int i = 0;
                        foreach (var anchorableToImport in layoutAnchorablePaneGroup.Descendents().OfType<LayoutAnchorable>().ToArray()) {
                            paneModel.Children.Insert(i, anchorableToImport);
                            i++;
                        }
                    }
 break;
                    #endregion
            }

            base.Drop(floatingWindow);
        }

        public override System.Windows.Media.Geometry GetPreviewPath(OverlayWindow overlayWindow, LayoutFloatingWindow floatingWindowModel) {
            switch (Type) {
                case DropTargetType.DocumentPaneGroupDockInside:
                    #region DropTargetType.DocumentPaneGroupDockInside
 {
                        var targetScreenRect = TargetElement.GetScreenArea();
                        targetScreenRect.Offset(-overlayWindow.Left, -overlayWindow.Top);

                        return new RectangleGeometry(targetScreenRect);
                    }
                    #endregion
            }

            return null;
        }
    }
}
