﻿/************************************************************************

   AvalonDock

   Copyright (C) 2007-2013 Xceed Software Inc.

   This program is provided to you under the terms of the New BSD
   License (BSD) as published at http://avalondock.codeplex.com/license 

   For more features, controls, and fast professional support,
   pick up AvalonDock in Extended WPF Toolkit Plus at http://xceed.com/wpf_toolkit

   Stay informed: follow @datagrid on Twitter or Like facebook.com/datagrids

  **********************************************************************/

using System.Windows;
using System.Windows.Controls;

namespace Xceed.Wpf.AvalonDock.Controls {
    public class MenuItemEx : MenuItem {
        static MenuItemEx() {
            IconProperty.OverrideMetadata(typeof(MenuItemEx), new FrameworkPropertyMetadata(new PropertyChangedCallback(OnIconPropertyChanged)));
        }


        public MenuItemEx() {
        }

        #region IconTemplate

        /// <summary>
        /// IconTemplate Dependency Property
        /// </summary>
        public static readonly DependencyProperty IconTemplateProperty =
            DependencyProperty.Register("IconTemplate", typeof(DataTemplate), typeof(MenuItemEx),
                new FrameworkPropertyMetadata((DataTemplate)null,
                    new PropertyChangedCallback(OnIconTemplateChanged)));

        /// <summary>
        /// Gets or sets the IconTemplate property.  This dependency property 
        /// indicates the data template for the icon.
        /// </summary>
        public DataTemplate IconTemplate {
            get { return (DataTemplate)GetValue(IconTemplateProperty); }
            set { SetValue(IconTemplateProperty, value); }
        }

        /// <summary>
        /// Handles changes to the IconTemplate property.
        /// </summary>
        private static void OnIconTemplateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) {
            ((MenuItemEx)d).OnIconTemplateChanged(e);
        }

        /// <summary>
        /// Provides derived classes an opportunity to handle changes to the IconTemplate property.
        /// </summary>
        protected virtual void OnIconTemplateChanged(DependencyPropertyChangedEventArgs e) {
            UpdateIcon();
        }

        #endregion

        #region IconTemplateSelector

        /// <summary>
        /// IconTemplateSelector Dependency Property
        /// </summary>
        public static readonly DependencyProperty IconTemplateSelectorProperty =
            DependencyProperty.Register("IconTemplateSelector", typeof(DataTemplateSelector), typeof(MenuItemEx),
                new FrameworkPropertyMetadata((DataTemplateSelector)null,
                    new PropertyChangedCallback(OnIconTemplateSelectorChanged)));

        /// <summary>
        /// Gets or sets the IconTemplateSelector property.  This dependency property 
        /// indicates the DataTemplateSelector for the Icon.
        /// </summary>
        public DataTemplateSelector IconTemplateSelector {
            get { return (DataTemplateSelector)GetValue(IconTemplateSelectorProperty); }
            set { SetValue(IconTemplateSelectorProperty, value); }
        }

        /// <summary>
        /// Handles changes to the IconTemplateSelector property.
        /// </summary>
        private static void OnIconTemplateSelectorChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) {
            ((MenuItemEx)d).OnIconTemplateSelectorChanged(e);
        }

        /// <summary>
        /// Provides derived classes an opportunity to handle changes to the IconTemplateSelector property.
        /// </summary>
        protected virtual void OnIconTemplateSelectorChanged(DependencyPropertyChangedEventArgs e) {
            UpdateIcon();
        }

        #endregion

        static void OnIconPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e) {
            if (e.NewValue != null) {
                ((MenuItemEx)sender).UpdateIcon();
            }
        }

        bool _reentrantFlag = false;
        void UpdateIcon() {
            if (_reentrantFlag)
                return;
            _reentrantFlag = true;
            if (IconTemplateSelector != null) {
                var dataTemplateToUse = IconTemplateSelector.SelectTemplate(Icon, this);
                if (dataTemplateToUse != null)
                    Icon = dataTemplateToUse.LoadContent();
            }
            else if (IconTemplate != null)
                Icon = IconTemplate.LoadContent();
            _reentrantFlag = false;
        }
    }
}
