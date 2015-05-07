/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.CatalogUI;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.NetworkAnalysis;
using Microsoft.Win32;
using Microsoft.Windows.Controls.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Xceed.Wpf.AvalonDock.Themes;

namespace ESRI.PrototypeLab.ZetaControls {
    public partial class GeometricNetworkWindow : RibbonWindow {
        private DispatcherTimer _timer = null;
        public const string DATAOBJECT_ESRINAMES = "ESRI Names";

        private const string GUID_SIMPLEJUNCTION_CLSID = "{CEE8D6B8-55FE-11D1-AE55-0000F80372B4}";
        private const string GUID_SIMPLEEDGE_CLSID = "{E7031C90-55FE-11D1-AE55-0000F80372B4}";
        private const string GUID_COMPLEXEDGE_CLSID = "{A30E8A2A-C50B-11D1-AEA9-0000F80372B4}";

        private SnapTolerance SNAP_XY = SnapTolerance.Minimum;
        private double SNAP_XY_CUSTOM = 0d;
        private bool PRESERVE_EXISTING_ENABLED_VALUES = true;
        private SnapTolerance SNAP_Z = SnapTolerance.None;
        private double SNAP_Z_CUSTOM = 0d;
        //
        // CONSTRUCTOR
        //
        public GeometricNetworkWindow() {
            InitializeComponent();

            // Handle drag events
            this.DragEnter += (s, e) => {
                e.Handled = true;
                if (GeometricNetworkViewModel.Default.Dataset != null) {
                    e.Effects = DragDropEffects.None;
                    return;
                }
                if (e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetDataPresent(GeometricNetworkWindow.DATAOBJECT_ESRINAMES)) {
                    e.Effects = DragDropEffects.All;
                    return;
                }

                e.Effects = DragDropEffects.None;
            };
            this.DragLeave += (s, e) => {
                e.Handled = true;
            };
            this.DragOver += (s, e) => {
                e.Handled = true;
                if (GeometricNetworkViewModel.Default.Dataset != null) {
                    e.Effects = DragDropEffects.None;
                    return;
                }
                if (e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetDataPresent(GeometricNetworkWindow.DATAOBJECT_ESRINAMES)) {
                    e.Effects = DragDropEffects.All;
                    return;
                }

                e.Effects = DragDropEffects.None;
            };
            this.Drop += new DragEventHandler(this.Window_Drop);

            // Start timer
            this._timer = new DispatcherTimer() {
                Interval = TimeSpan.FromMilliseconds(300d)
            };
            this._timer.Tick += this.Timer_Tick;
            this._timer.Start();

            // Perform actions on window load
            this.Loaded += (s, e) => {
                this.DockableContentOutput.ToggleAutoHide();
            };

            // Button click events
            this.ButtonOpen.Click += this.Button_Click;
            this.ButtonSave.Click += this.Button_Click;
            this.ButtonSaveAs.Click += this.Button_Click;
            this.ButtonClose.Click += this.Button_Click;
            this.ButtonImport.Click += this.Button_Click;
            this.ButtonExport.Click += this.Button_Click;
            this.ButtonOutput.Click += this.Button_Click;
            this.ButtonClassAdd.Click += this.Button_Click;
            this.ButtonClassRemove.Click += this.Button_Click;
            this.ButtonWeightAdd.Click += this.Button_Click;
            this.ButtonWeightRemove.Click += this.Button_Click;
            this.ButtonJunctionAdd.Click += this.Button_Click;
            this.ButtonJunctionRemove.Click += this.Button_Click;
            this.ButtonJunctionSetAsDefault.Checked += (s, e) => {
                // Exit if invalid
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }
                if (GeometricNetworkViewModel.Default.SelectedJunctionRule == null) { return; }

                //
                var r = GeometricNetworkViewModel.Default.SelectedJunctionRule;
                if (r.IsDefault) { return; }

                r.IsDefault = true;
                ZGeometricNetwork gn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
                var rules = gn.JunctionRules
                                    .Select(j => { return (ZJunctionConnectivityRule)j; })
                                    .Where(j => j != r)
                                    .Where(j => j.Edge == r.Edge);
                foreach (var rule in rules) {
                    if (rule.IsDefault) {
                        rule.IsDefault = false;
                    }
                }
                
                // Make document dirty
                GeometricNetworkViewModel.Default.MakeDirty();
            };
            this.ButtonJunctionSetAsDefault.Unchecked += (s, e) => {
                // Exit if invalid
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }
                if (GeometricNetworkViewModel.Default.SelectedJunctionRule == null) { return; }

                //
                var r = GeometricNetworkViewModel.Default.SelectedJunctionRule;
                if (!r.IsDefault) { return; }

                r.IsDefault = false;

                // Make document dirty
                GeometricNetworkViewModel.Default.MakeDirty();
            };

            this.MenuItemExit.Click += this.Button_Click;

            this.ButtonEdgeAddJunction.DropDownOpened += this.Button_DropDownOpened;
            this.ButtonEdgeRemoveJunction.DropDownOpened += this.Button_DropDownOpened;
            this.ButtonEdgeDefaultJunction.DropDownOpened += this.Button_DropDownOpened;

            this.RibbonGalleryEdgeAddJunction.SelectionChanged += this.RibbonGallery_SelectionChanged;
            this.RibbonGalleryEdgeRemoveJunction.SelectionChanged += this.RibbonGallery_SelectionChanged;
            this.RibbonGalleryEdgeDefaultJunction.SelectionChanged += this.RibbonGallery_SelectionChanged;

            GeometricNetworkViewModel.Default.Messages.CollectionChanged += (s, e) => {
                if (e.Action == NotifyCollectionChangedAction.Add) {
                    this.ItemsControlMessage.ScrollIntoView();
                }
            };

            // Network classes datagrid
            this.DataGridNetworkClasses.SelectionChanged += this.DataGrid_SelectionChanged;

            // Weights datagrid
            this.DataGridWeights.SelectionChanged += this.DataGrid_SelectionChanged;

            // Junction rule datagrid
            this.DataGridJunctionRules.SelectedCellsChanged += this.DataGrid_SelectedCellsChanged;
            this.DataGridJunctionRules.AutoGeneratingColumn += this.DataGrid_AutoGeneratingColumn;

            // Edge rule datagrid
            this.DataGridEdgeRules.SelectedCellsChanged += this.DataGrid_SelectedCellsChanged;
            this.DataGridEdgeRules.AutoGeneratingColumn += this.DataGrid_AutoGeneratingColumn;

            // Change docking environment visual theme
            this.dockManager.Theme = new AeroTheme();

            // Assign view model as data context
            this.DataContext = GeometricNetworkViewModel.Default;
        }
        private void RibbonGallery_SelectionChanged(object sender, RoutedPropertyChangedEventArgs<object> e) {
            var rule = GeometricNetworkViewModel.Default.SelectedEdgeRule;
            ZSubtype s = e.NewValue as ZSubtype;
            if (s == null){return;}

            if (sender == this.RibbonGalleryEdgeAddJunction) {
                if (!rule.Junctions.Contains(s)) {
                    rule.Junctions.Add(s);
                    // Make document dirty
                    GeometricNetworkViewModel.Default.MakeDirty();
                }
            }
            else if (sender == this.RibbonGalleryEdgeRemoveJunction){
                if (rule.Junctions.Contains(s)) {
                    rule.Junctions.Remove(s);

                    // Make document dirty
                    GeometricNetworkViewModel.Default.MakeDirty();
                }
            }
            else if (sender == this.RibbonGalleryEdgeDefaultJunction) {
                if (rule.DefaultJunction != s) {
                    rule.DefaultJunction = s;

                    // Make document dirty
                    GeometricNetworkViewModel.Default.MakeDirty();
                }
            }
        }
        private void Button_DropDownOpened(object sender, EventArgs e) {
            // Exit if geometric network is missing
            var gn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
            if (gn == null){return;}

            // Exit if no rule selected
            var rule = GeometricNetworkViewModel.Default.SelectedEdgeRule;
            if (rule == null) { return; }

            if (sender == this.ButtonEdgeAddJunction) {
                var list = new List<ZSubtype>();
                foreach (ZNetworkClass nc in gn.NetworkClasses.Where(n => n.IsJunction).OrderBy(n => n.Path.Table)) {
                    foreach (ZSubtype st in nc.Subtypes.OrderBy(s => s.Name)) {
                        // Add row to list
                        list.Add(st);
                    }
                }
                var add = list.Except(rule.Junctions);
                this.RibbonGalleryCategoryEdgeAddJunction.ItemsSource = add;
            }
            else if (sender == this.ButtonEdgeRemoveJunction) {
                this.RibbonGalleryCategoryEdgeRemoveJunction.ItemsSource = rule.Junctions;
            }
            else if (sender == this.ButtonEdgeDefaultJunction) {
                this.RibbonGalleryCategoryEdgeDefaultJunction.ItemsSource = rule.Junctions;
                this.RibbonGalleryEdgeDefaultJunction.SelectedItem = rule.DefaultJunction;
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e) {
            if (sender == this.MenuItemExit) {
                if (GeometricNetworkViewModel.Default.IsDirty) {
                    MessageBoxResult r = MessageBox.Show(
                        "Do you want to save before closing?",
                        GeometricNetworkViewModel.Default.WindowTitle,
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Exclamation,
                        MessageBoxResult.Yes
                    );
                    switch (r) {
                        case MessageBoxResult.Yes:
                            // Save document
                            GeometricNetworkViewModel.Default.Save();

                            // If user canceled save dialog then exit
                            if (GeometricNetworkViewModel.Default.IsDirty) { return; }

                            break;
                        case MessageBoxResult.No:
                            break;
                        case MessageBoxResult.Cancel:
                        default:
                            return;
                    }
                }

                // Exit application
                Application.Current.Shutdown(0);
            }
            else if (sender == this.ButtonOpen) {
                // Exit if dataset already open
                if (GeometricNetworkViewModel.Default.Dataset != null) {
                    MessageBox.Show(
                        "Please close the document first",
                        GeometricNetworkViewModel.Default.WindowTitle,
                        MessageBoxButton.OK,
                        MessageBoxImage.Information,
                        MessageBoxResult.OK
                    );
                    return;
                }

                OpenFileDialog openFileDialog = new OpenFileDialog() {
                    CheckFileExists = true,
                    Filter = "GN definition document" + " (*.esriGeoNet)|*.esriGeoNet",
                    FilterIndex = 1,
                    Multiselect = false,
                    Title = GeometricNetworkViewModel.Default.WindowTitle
                };

                // Check if user pressed "Save" and File is OK.
                bool? ok = openFileDialog.ShowDialog(this);
                if (!ok.HasValue || !ok.Value) { return; }
                if (string.IsNullOrWhiteSpace(openFileDialog.FileName)) { return; }

                //
                GeometricNetworkViewModel.Default.Load(openFileDialog.FileName);
            }
            else if (sender == this.ButtonSave) {
                // Handle event to prevent bubbling event to ButtonSave
                if (e != null) {
                    e.Handled = true;
                }

                // Exit if no dataset
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }

                // If no document, show "save as" dialog
                if (GeometricNetworkViewModel.Default.Document == null) {
                    this.Button_Click(this.ButtonSaveAs, null);
                    return;
                }

                // Save document
                GeometricNetworkViewModel.Default.Save();
            }
            else if (sender == this.ButtonSaveAs) {
                // Handle event to prevent bubbling event to ButtonSave
                if (e != null) {
                    e.Handled = true;
                }

                // Show save dialog
                SaveFileDialog saveFileDialog = new SaveFileDialog() {
                    DefaultExt = "esriGeoNet",
                    FileName = "Document1",
                    Filter = "GN definition document" + " (*.esriGeoNet)|*.esriGeoNet",
                    FilterIndex = 1,
                    OverwritePrompt = true,
                    RestoreDirectory = false,
                    Title = GeometricNetworkViewModel.Default.WindowTitle
                };

                // Check if user pressed "Save" and File is OK.
                bool? ok = saveFileDialog.ShowDialog(this);
                if (!ok.HasValue || !ok.Value) { return; }
                if (string.IsNullOrWhiteSpace(saveFileDialog.FileName)) { return; }

                //
                GeometricNetworkViewModel.Default.Save(saveFileDialog.FileName);
            }
            else if (sender == this.ButtonClose) {
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }
                if (GeometricNetworkViewModel.Default.IsDirty) {
                    MessageBoxResult r = MessageBox.Show(
                        "Do you want to save before closing?",
                        GeometricNetworkViewModel.Default.WindowTitle,
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Exclamation,
                        MessageBoxResult.Yes
                    );
                    switch (r) {
                        case MessageBoxResult.Yes:
                            GeometricNetworkViewModel.Default.Save();
                            break;
                        case MessageBoxResult.No:
                            break;
                        case MessageBoxResult.Cancel:
                        default:
                            return;
                    }
                }

                // Clear current dataset
                GeometricNetworkViewModel.Default.Clear();
            }
            else if (sender == this.ButtonImport) {
                // Create GxObjectFilter for GxDialog
                IGxObjectFilter gxObjectFilter = new GxFilterGeometricNetworksClass();

                // Create GxDialog
                IGxDialog gxDialog = new GxDialogClass() {
                    AllowMultiSelect = false,
                    ButtonCaption = "Import",
                    ObjectFilter = gxObjectFilter,
                    RememberLocation = true,
                    Title = "Please select a geometric network"
                };

                // Declare Enumerator to hold selected objects
                IEnumGxObject enumGxObject = null;

                // Open Dialog
                if (!gxDialog.DoModalOpen(0, out enumGxObject)) { return; }
                if (enumGxObject == null) { return; }

                // Get Selected Object (if any)
                IGxObject gxObject = enumGxObject.Next();
                if (gxObject == null) { return; }
                if (!gxObject.IsValid) { return; }

                // Get GxDataset
                if (!(gxObject is IGxDataset)) { return; }
                IGxDataset gxDataset = (IGxDataset)gxObject;

                // Load geometric network from named object
                IName name = (IName)gxDataset.DatasetName;
                GeometricNetworkLoader loader = new GeometricNetworkLoader(name);
                loader.Load();
            }
            else if (sender == this.ButtonExport) {
                ResultType ok = this.ExportGeometricNetwork();
                switch (ok) {
                    case ResultType.Cancelled:
                        break;
                    case ResultType.Error:
                        MessageBox.Show(
                            "Geometric network creation failed",
                            GeometricNetworkViewModel.Default.WindowTitle,
                            MessageBoxButton.OK,
                            MessageBoxImage.Information,
                            MessageBoxResult.OK
                        );
                        break;
                    case ResultType.Successful:
                        MessageBox.Show(
                           "Geometric network creation successful",
                           GeometricNetworkViewModel.Default.WindowTitle,
                           MessageBoxButton.OK,
                           MessageBoxImage.Information,
                           MessageBoxResult.OK
                       );
                        break;
                }
            }
            else if (sender == this.ButtonOutput) {
                if (this.DockableContentOutput.IsHidden) {
                    this.DockableContentOutput.Show();
                }
                else if (this.DockableContentOutput.IsAutoHidden) {
                    this.DockableContentOutput.ToggleAutoHide();
                }
            }
            else if (sender == this.ButtonJunctionAdd) {
                // Exit if invalid
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }
                if (GeometricNetworkViewModel.Default.SelectedJunctionRule != null) { return; }

                // Add junction rule
                if (this.DataGridJunctionRules.SelectedCells == null) { return; }
                if (this.DataGridJunctionRules.SelectedCells.Count == 0) { return; }
                DataGridCellInfo ci = this.DataGridJunctionRules.SelectedCells[0];
                object o = ci.Item;
                if (o == null) { return; }

                RuleDataGridColumn rdgc = ci.Column as RuleDataGridColumn;
                if (rdgc == null) { return; }

                ZGeometricNetwork zgn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
                if (zgn == null) { return; }

                // Create new rule
                ZSubtype e1 = o.GetType().GetProperty(RuleMatrix.EDGE_SUBTYPE).GetValue(o, null) as ZSubtype;
                ZSubtype j1 = zgn.FindSubtype(rdgc.ColumnName);
                ZJunctionConnectivityRule rule = new ZJunctionConnectivityRule(e1, j1);

                // Add rule to network
                zgn.JunctionRules.Add(rule);

                // Update data source
                o.GetType().GetProperty(rdgc.ColumnName).SetValue(o, rule, null);

                // Refresh display
                GeometricNetworkViewModel.Default.SelectedJunctionRule = (ZJunctionConnectivityRule)rule;
                GeometricNetworkViewModel.Default.JunctionRuleDataSource.Refresh();

                // Focus cell
                DataGridCellInfo cell = new DataGridCellInfo(o, rdgc);
                this.DataGridJunctionRules.CurrentCell = cell;
                this.DataGridJunctionRules.SelectedCells.Clear();
                this.DataGridJunctionRules.SelectedCells.Add(cell);

                // Make document dirty
                GeometricNetworkViewModel.Default.MakeDirty();
            }
            else if (sender == this.ButtonJunctionRemove) {
                // Exit if invalid
                if (GeometricNetworkViewModel.Default.Dataset == null) { return; }
                if (GeometricNetworkViewModel.Default.SelectedJunctionRule == null) { return; }

                // Get selected cell
                if (this.DataGridJunctionRules.SelectedCells == null) { return; }
                if (this.DataGridJunctionRules.SelectedCells.Count == 0) { return; }
                DataGridCellInfo ci = this.DataGridJunctionRules.SelectedCells[0];
                object o = ci.Item;
                if (o == null) { return; }

                // Get selected rule
                RuleDataGridColumn rdgc = ci.Column as RuleDataGridColumn;
                if (rdgc == null) { return; }
                ZRule rule = o.GetType().GetProperty(rdgc.ColumnName).GetValue(o, null) as ZRule;
                if (rule == null) { return; }

                // Update data source
                o.GetType().GetProperty(rdgc.ColumnName).SetValue(o, null, null);

                // Remove from dataset
                ZGeometricNetwork zgn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
                if (zgn == null) { return; }
                zgn.JunctionRules.Remove(rule);

                // Refresh display
                GeometricNetworkViewModel.Default.SelectedJunctionRule = (ZJunctionConnectivityRule)rule;
                GeometricNetworkViewModel.Default.JunctionRuleDataSource.Refresh();

                // Focus cell
                DataGridCellInfo cell = new DataGridCellInfo(o, rdgc);
                this.DataGridJunctionRules.CurrentCell = cell;
                this.DataGridJunctionRules.SelectedCells.Clear();
                this.DataGridJunctionRules.SelectedCells.Add(cell);

                // Make document dirty
                GeometricNetworkViewModel.Default.MakeDirty();
            }
        }
        private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e) {
            // Nothing selected?
            if (e.AddedCells.Count == 0) { return; }

            // Prevent user from selecting the first "subtype" column
            SubtypeDataGridColumn colum = e.AddedCells[0].Column as SubtypeDataGridColumn;
            if (colum != null) {
                if (e.RemovedCells.Count != 0) {
                    DataGrid dg = sender as DataGrid;
                    DataGridCellInfo ci = e.RemovedCells[0];
                    DataGridCellInfo cell = new DataGridCellInfo(ci.Item, ci.Column);
                    dg.CurrentCell = cell;
                    dg.SelectedCells.Clear();
                    dg.SelectedCells.Add(cell);
                    return;
                }
            }

            // Update the selected junction/edge dependancy property
            if (sender == this.DataGridJunctionRules) {
                GeometricNetworkViewModel.Default.SelectedJunctionRule = this.GetSelectedRule(this.DataGridJunctionRules) as ZJunctionConnectivityRule;
            }
            else if (sender == this.DataGridEdgeRules) {
                GeometricNetworkViewModel.Default.SelectedEdgeRule = this.GetSelectedRule(this.DataGridEdgeRules) as ZEdgeConnectivityRule;
            }
        }
        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (sender == this.DataGridNetworkClasses) {
                if (e == null || e.AddedItems == null || e.AddedItems.Count == 0) {
                    GeometricNetworkViewModel.Default.SelectedNetworkClass = null;
                    return;
                }
                ZNetworkClass n = e.AddedItems[0] as ZNetworkClass;
                GeometricNetworkViewModel.Default.SelectedNetworkClass = n;
            }
            else if (sender == this.DataGridWeights) {
                if (e == null || e.AddedItems == null || e.AddedItems.Count == 0) {
                    GeometricNetworkViewModel.Default.SelectedWeight = null;
                    return;
                }
                ZNetWeight n = e.AddedItems[0] as ZNetWeight;
                GeometricNetworkViewModel.Default.SelectedWeight = n;
            }
        }
        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e) {
            //
            DataGridTemplateColumn dataGridColumn = null;

            if (sender == this.DataGridJunctionRules) {
                switch (e.PropertyName) {
                    case RuleMatrix.EDGE_FEATURECLASS:
                        break;
                    case RuleMatrix.EDGE_SUBTYPE:
                        dataGridColumn = new SubtypeDataGridColumn() {
                            Header = string.Empty,
                            CellTemplate = this.Resources["SubtypeCellTemplate"] as DataTemplate
                        };
                        break;
                    default:
                        // Get subtype by guid
                        ZGeometricNetwork zgn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
                        ZSubtype subtype = zgn.FindSubtype(e.PropertyName);

                        // Create data column
                        dataGridColumn = new RuleDataGridColumn() {
                            Header = subtype,
                            HeaderTemplate = this.Resources["JunctionRuleFeatureclassHeaderTemplate"] as DataTemplate,
                            CellTemplate = this.Resources["JunctionRuleCellTemplate"] as DataTemplate,
                            ColumnName = e.PropertyName
                        };

                        break;
                }
            }
            else if (sender == this.DataGridEdgeRules) {
                switch (e.PropertyName) {
                    case RuleMatrix.EDGE_FEATURECLASS:
                        break;
                    case RuleMatrix.EDGE_SUBTYPE:
                        dataGridColumn = new SubtypeDataGridColumn() {
                            Header = string.Empty,
                            CellTemplate = this.Resources["SubtypeCellTemplate"] as DataTemplate
                        };
                        break;
                    default:
                        // Get subtype by guid
                        ZGeometricNetwork zgn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
                        ZSubtype subtype = zgn.FindSubtype(e.PropertyName);

                        // Create data column
                        dataGridColumn = new RuleDataGridColumn() {
                            Header = subtype,
                            HeaderTemplate = this.Resources["EdgeRuleFeatureclassHeaderTemplate"] as DataTemplate,
                            HeaderStyle = this.Resources["EdgeRuleHeaderStyle"] as Style,
                            CellTemplate = this.Resources["EdgeRuleCellTemplate"] as DataTemplate,
                            ColumnName = e.PropertyName
                        };

                        break;
                }
            }

            //
            e.Column = dataGridColumn;
        }
        private void DataGrid_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            var dataGrid = (DataGrid)sender;
            if (dataGrid.IsGrouping && e.HorizontalChange != 0d) {
                GeometricNetworkWindow.TraverseVisualTree(dataGrid, e.HorizontalOffset);
            }
        }
        private void Window_Drop(object sender, DragEventArgs e) {
            try {
                // Handle event
                e.Handled = true;

                DataObject d = e.Data as DataObject;

                if (d.GetDataPresent(DataFormats.FileDrop)) {
                    if (d.ContainsFileDropList()) {
                        StringCollection sc = d.GetFileDropList();
                        string file = sc[0];
                        GeometricNetworkViewModel.Default.Load(file);
                        GeometricNetworkViewModel.Default.CreateRuleDataSource();
                        GeometricNetworkViewModel.Default.Document = file;
                    }
                }
                else if (d.GetDataPresent(GeometricNetworkWindow.DATAOBJECT_ESRINAMES)) {
                    // Get Dropped Object
                    object objectDrop = d.GetData(GeometricNetworkWindow.DATAOBJECT_ESRINAMES);
                    MemoryStream memoryStream = (MemoryStream)objectDrop;
                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();

                    // Create loader
                    GeometricNetworkLoader loader = new GeometricNetworkLoader(bytes);
                    loader.Load();
                }
            }
            catch (Exception ex) {
#if DEBUG
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
#endif
            }
        }
        private ZRule GetSelectedRule(DataGrid grid) {
            if (grid.SelectedCells == null) { return null; }
            if (grid.SelectedCells.Count == 0) { return null; }
            DataGridCellInfo ci = grid.SelectedCells[0];
            object o = ci.Item;
            if (o == null) { return null; }

            RuleDataGridColumn rdgc = ci.Column as RuleDataGridColumn;
            if (rdgc == null) { return null; }

            ZRule rule = o.GetType().GetProperty(rdgc.ColumnName).GetValue(o, null) as ZRule;
            return rule;
        }
        private void Timer_Tick(object sender, EventArgs e) {
            if (!this.IsLoaded) { return; }
            if (!this.IsVisible) { return; }

            bool isOpen = GeometricNetworkViewModel.Default.Dataset != null;

            switch (this.Ribbon.SelectedIndex) {
                case 0: // Home tab
                    // Open
                    if (this.ButtonOpen.IsEnabled != !isOpen) {
                        this.ButtonOpen.IsEnabled = !isOpen;
                    }

                    // Save
                    bool save = isOpen && GeometricNetworkViewModel.Default.IsDirty;
                    if (this.ButtonSave.IsEnabled != save) {
                        this.ButtonSave.IsEnabled = save;
                    }

                    // Save As
                    bool saveas = isOpen;
                    if (this.ButtonSaveAs.IsEnabled != saveas) {
                        this.ButtonSaveAs.IsEnabled = saveas;
                    }

                    // Close
                    if (this.ButtonClose.IsEnabled != isOpen) {
                        this.ButtonClose.IsEnabled = isOpen;
                    }

                    // Import
                    if (this.ButtonImport.IsEnabled != !isOpen) {
                        this.ButtonImport.IsEnabled = !isOpen;
                    }

                    // Export
                    if (this.ButtonExport.IsEnabled != isOpen) {
                        this.ButtonExport.IsEnabled = isOpen;
                    }

                    // Output
                    bool output = isOpen && (this.DockableContentOutput.IsHidden || this.DockableContentOutput.IsAutoHidden);
                    if (this.ButtonOutput.IsEnabled != output) {
                        this.ButtonOutput.IsEnabled = output;
                    }

                    break;
                case 1: // Editor tab
                    // Add class
                    bool addclass = false; // isOpen && this.LayoutDocumentClasses.IsSelected;
                    if (this.ButtonClassAdd.IsEnabled != addclass) {
                        this.ButtonClassAdd.IsEnabled = addclass;
                    }

                    // Remove class
                    bool removeclass = false; // isOpen && this.LayoutDocumentClasses.IsSelected && GeometricNetworkViewModel.Default.SelectedNetworkClass != null;
                    if (this.ButtonClassRemove.IsEnabled != removeclass) {
                        this.ButtonClassRemove.IsEnabled = removeclass;
                    }

                    // Add weight
                    bool addweight = false; //isOpen && this.LayoutDocumentWeights.IsSelected;
                    if (this.ButtonWeightAdd.IsEnabled != addweight) {
                        this.ButtonWeightAdd.IsEnabled = addweight;
                    }

                    // Remove weight
                    bool removeweight = false; //isOpen && this.LayoutDocumentWeights.IsSelected && GeometricNetworkViewModel.Default.SelectedWeight != null;
                    if (this.ButtonWeightRemove.IsEnabled != removeweight) {
                        this.ButtonWeightRemove.IsEnabled = removeweight;
                    }

                    // Add junction
                    bool addjunction = isOpen && this.LayoutDocumentJunctionRules.IsSelected && GeometricNetworkViewModel.Default.SelectedJunctionRule == null;
                    if (this.ButtonJunctionAdd.IsEnabled != addjunction) {
                        this.ButtonJunctionAdd.IsEnabled = addjunction;
                    }

                    // Remove junction
                    bool removejunction = isOpen && this.LayoutDocumentJunctionRules.IsSelected && GeometricNetworkViewModel.Default.SelectedJunctionRule != null;
                    if (this.ButtonJunctionRemove.IsEnabled != removejunction) {
                        this.ButtonJunctionRemove.IsEnabled = removejunction;
                    }

                    // Min edge / max edge / min junction / max junction
                    bool junctionrulecardinality = isOpen && this.LayoutDocumentJunctionRules.IsSelected && GeometricNetworkViewModel.Default.SelectedJunctionRule != null;
                    if (this.GridJunctionRuleCardinality.IsEnabled != junctionrulecardinality) {
                        this.GridJunctionRuleCardinality.IsEnabled = junctionrulecardinality;
                    }

                    // Set as default junction
                    bool junctionsetasdefault = isOpen && this.LayoutDocumentJunctionRules.IsSelected && GeometricNetworkViewModel.Default.SelectedJunctionRule != null;
                    if (this.ButtonJunctionSetAsDefault.IsEnabled != junctionsetasdefault) {
                        this.ButtonJunctionSetAsDefault.IsEnabled = junctionsetasdefault;
                    }
                    bool isdefault = isOpen && GeometricNetworkViewModel.Default.SelectedJunctionRule != null && GeometricNetworkViewModel.Default.SelectedJunctionRule.IsDefault;
                    if (this.ButtonJunctionSetAsDefault.IsChecked != isdefault) {
                        this.ButtonJunctionSetAsDefault.IsChecked = isdefault;
                    };

                    // Add junction to EE rule
                    bool addjunction2 = isOpen && this.LayoutDocumentEdgeRules.IsSelected && GeometricNetworkViewModel.Default.SelectedEdgeRule != null;
                    if (this.ButtonEdgeAddJunction.IsEnabled != addjunction2) {
                        this.ButtonEdgeAddJunction.IsEnabled = addjunction2;
                    }

                    // Remove junction to EE rule
                    bool removejunction2 = isOpen && this.LayoutDocumentEdgeRules.IsSelected && GeometricNetworkViewModel.Default.SelectedEdgeRule != null;
                    if (this.ButtonEdgeRemoveJunction.IsEnabled != removejunction2) {
                        this.ButtonEdgeRemoveJunction.IsEnabled = removejunction2;
                    }

                    // Default junction of edge rule
                    bool defaultjunction = isOpen && this.LayoutDocumentEdgeRules.IsSelected && GeometricNetworkViewModel.Default.SelectedEdgeRule != null;
                    if (this.ButtonEdgeDefaultJunction.IsEnabled != defaultjunction) {
                        this.ButtonEdgeDefaultJunction.IsEnabled = defaultjunction;
                    }

                    break;
                case 2: // Review tab
                    break;

            }

            // Document
            Visibility v = isOpen ? Visibility.Visible : Visibility.Collapsed;
            if (this.dockManager.Visibility != v) {
                this.dockManager.Visibility = v;
            }
        }
        private ResultType ExportGeometricNetwork() {
            // Create FeatureDataset filter for GxDialog
            IGxObjectFilter gxObjectFilter = new GxFilterGeometricNetworksClass();

            // Create GxDialog
            IGxDialog gxDialog = new GxDialogClass() {
                AllowMultiSelect = false,
                ButtonCaption = "Export",
                ObjectFilter = gxObjectFilter,
                RememberLocation = true,
                Title = "Save Geometric Network As"
            };

            // Open Dialog
            if (!gxDialog.DoModalSave(0)) { return ResultType.Cancelled; }

            // Must specify a name
            if (string.IsNullOrWhiteSpace(gxDialog.Name)) {
                GeometricNetworkViewModel.Default.AddMessage("You must specify a name", MessageType.Error);
                return ResultType.Error;
            }

            // Check parent feature dataset
            IGxObject fd = gxDialog.FinalLocation;
            if (fd == null || !fd.IsValid) {
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Invalid feature dataset", gxDialog.Name), MessageType.Error);
                return ResultType.Error;
            }

            // Get feature dataset and workspace
            IGxDataset dataset = (IGxDataset)fd;
            IFeatureDataset d = (IFeatureDataset)dataset.Dataset;
            IWorkspace w = d.Workspace;

            // Check if geometric network already exists
            IGeometricNetwork gn = w.FindGeometricNetwork(gxDialog.Name);
            if (gn != null) {
                MessageBoxResult r = MessageBox.Show(
                    "The geometric network already exists. Remove?",
                    GeometricNetworkViewModel.Default.WindowTitle,
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Exclamation,
                    MessageBoxResult.Yes
                );
                switch (r) {
                    case MessageBoxResult.Yes:
                        IDataset dd = (IDataset)gn;

                        if (!dd.CanDelete()) {
                            GeometricNetworkViewModel.Default.AddMessage("Cannot delete geometric network", MessageType.Error);
                            return ResultType.Error;
                        }

                        try {
                            dd.Delete();
                            GeometricNetworkViewModel.Default.AddMessage("Deletion successful", MessageType.Information);
                        }
                        catch (Exception ex) {
                            GeometricNetworkViewModel.Default.AddMessage(ex.Message, MessageType.Error);
                            return ResultType.Error;
                        }

                        break;
                    case MessageBoxResult.No:
                    case MessageBoxResult.Cancel:
                    default:
                        return ResultType.Cancelled;
                }
            }

            // Initialize loader
            INetworkLoader3 networkLoader = new NetworkLoaderClass() {
                FeatureDatasetName = (IDatasetName)d.FullName,
                NetworkName = gxDialog.Name,
                NetworkType = esriNetworkType.esriNTUtilityNetwork,
                PreserveEnabledValues = PRESERVE_EXISTING_ENABLED_VALUES
            };
            INetworkLoaderProps networkLoadedProps = (INetworkLoaderProps)networkLoader;
            INetworkLoaderProgress_Event ne = (INetworkLoaderProgress_Event)networkLoader;
            ne.PutMessage += (a, b) => {
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Progress: '{0}'", b), MessageType.Information);
            };

            // Assign XY snapping tolerance
            switch (SNAP_XY) {
                case SnapTolerance.None:
                    networkLoader.SnapTolerance = 0d;
                    networkLoader.UseXYsForSnapping = false;
                    break;
                case SnapTolerance.Default:
                    networkLoader.SnapTolerance = networkLoader.DefaultSnapTolerance;
                    networkLoader.UseXYsForSnapping = true;
                    break;
                case SnapTolerance.Minimum:
                    networkLoader.SnapTolerance = networkLoader.MinSnapTolerance;
                    networkLoader.UseXYsForSnapping = true;
                    break;
                case SnapTolerance.Maximum:
                    networkLoader.SnapTolerance = networkLoader.MaxSnapTolerance;
                    networkLoader.UseXYsForSnapping = true;
                    break;
                case SnapTolerance.Custom:
                    if (SNAP_XY_CUSTOM < networkLoader.MinSnapTolerance) {
                        networkLoader.SnapTolerance = networkLoader.MinSnapTolerance;
                    }
                    else if (SNAP_XY_CUSTOM > networkLoader.MaxSnapTolerance) {
                        networkLoader.SnapTolerance = networkLoader.MaxSnapTolerance;
                    }
                    else {
                        networkLoader.SnapTolerance = SNAP_XY_CUSTOM;
                    }
                    networkLoader.UseXYsForSnapping = true;
                    break;
            }

            // Assign Z snapping tolerance
            if (networkLoader.CanUseZs) {
                switch (SNAP_Z) {
                    case SnapTolerance.None:
                        networkLoader.ZSnapTolerance = 0d;
                        networkLoader.UseZs = false;
                        break;
                    case SnapTolerance.Default:
                        networkLoader.ZSnapTolerance = networkLoader.DefaultZSnapTolerance;
                        networkLoader.UseZs = true;
                        break;
                    case SnapTolerance.Minimum:
                        networkLoader.ZSnapTolerance = networkLoader.MinZSnapTolerance;
                        networkLoader.UseZs = true;
                        break;
                    case SnapTolerance.Maximum:
                        networkLoader.ZSnapTolerance = networkLoader.MaxZSnapTolerance;
                        networkLoader.UseZs = true;
                        break;
                    case SnapTolerance.Custom:
                        if (SNAP_Z_CUSTOM < networkLoader.MinZSnapTolerance) {
                            networkLoader.SnapTolerance = networkLoader.MinZSnapTolerance;
                        }
                        else if (SNAP_XY_CUSTOM > networkLoader.MaxZSnapTolerance) {
                            networkLoader.SnapTolerance = networkLoader.MaxZSnapTolerance;
                        }
                        else {
                            networkLoader.SnapTolerance = SNAP_Z_CUSTOM;
                        }
                        networkLoader.UseZs = true;
                        break;
                }
            }
            else {
                switch (SNAP_Z) {
                    case SnapTolerance.None:
                        break;
                    case SnapTolerance.Minimum:
                    case SnapTolerance.Maximum:
                    case SnapTolerance.Custom:
                        GeometricNetworkViewModel.Default.AddMessage("Z snapping is unavailable", MessageType.Warning);
                        break;
                }
            }

            // Check each feature class
            bool badfeatureclass = false;
            IFeatureClassContainer fcc = (IFeatureClassContainer)d;
            ZGeometricNetwork zgn = GeometricNetworkViewModel.Default.Dataset as ZGeometricNetwork;
            foreach (ZNetworkClass znc in zgn.NetworkClasses.Where(f => !f.IsOrphanJunctionFeatureClass)) {
                // Does the feature class exist?
                IFeatureClass fc = null;
                try {
                    fc = fcc.get_ClassByName(znc.Path.Table);
                }
                catch { }
                if (fc == null) {
                    GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' does not exists", znc.Path.Table), MessageType.Error);
                    badfeatureclass = true;
                    continue;
                }

                // Is the feature class compatiable?
                switch (networkLoader.CanUseFeatureClass(znc.Path.Table)) {
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCCannotOpen:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' cannot be opened", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCInAnotherNetwork:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' belongs to another network", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCInTerrain:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' belongs to a terrain", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCInTopology:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' belongs to a topology", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCInvalidFeatureType:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' has an invalid feature type", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCInvalidShapeType:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' has an invalid shape type", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCIsCompressedReadOnly:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' is compressed readonly", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCRegisteredAsVersioned:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' is registered as versioned", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCUnknownError:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("Feature class '{0}' has an unknown error", znc.Path.Table), MessageType.Error);
                        badfeatureclass = true;
                        break;
                    case esriNetworkLoaderFeatureClassCheck.esriNLFCCValid:
                        break;
                }
            }
            if (badfeatureclass) { return ResultType.Error; }

            // Add feature class to network
            bool snap = SNAP_XY == SnapTolerance.None;
            foreach (ZNetworkClass znc in zgn.NetworkClasses.Where(f => !f.IsOrphanJunctionFeatureClass)) {
                string table = znc.Path.Table;

                if (znc.IsJunction) {
                    // Add simple junction
                    UID uid = new UIDClass() {
                        Value = GUID_SIMPLEJUNCTION_CLSID
                    };
                    networkLoader.AddFeatureClass(table, esriFeatureType.esriFTSimpleJunction, uid, snap);

                    if (znc.IsSourceSink) {
                        // Check junction already has a role field
                        string field = networkLoadedProps.DefaultAncillaryRoleField;
                        switch (networkLoader.CheckAncillaryRoleField(table, field)) {
                            case esriNetworkLoaderFieldCheck.esriNLFCInvalidDomain:
                                GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - The AncillaryRole field has invalid domain", table, field), MessageType.Error);
                                return ResultType.Error;
                            case esriNetworkLoaderFieldCheck.esriNLFCInvalidType:
                                GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - The AncillaryRole field has invalid type", table, field), MessageType.Error);
                                return ResultType.Error;
                            case esriNetworkLoaderFieldCheck.esriNLFCNotFound:
                                GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - Adding AncillaryRole field", table, field), MessageType.Information);
                                networkLoader.PutAncillaryRole(table, esriNetworkClassAncillaryRole.esriNCARSourceSink, field);
                                break;
                            case esriNetworkLoaderFieldCheck.esriNLFCUnknownError:
                                GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - An unknown error was encountered", table, field), MessageType.Error);
                                return ResultType.Error;
                            case esriNetworkLoaderFieldCheck.esriNLFCValid:
                                GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - An AncillaryRole field already exists", table, field), MessageType.Warning);
                                networkLoader.PutAncillaryRole(table, esriNetworkClassAncillaryRole.esriNCARSourceSink, field);
                                break;
                        }
                    }
                }
                if (znc.IsEdge) {
                    if (znc.IsComplex) {
                        // Add simple junction
                        UID uid = new UIDClass() {
                            Value = GUID_COMPLEXEDGE_CLSID
                        };
                        networkLoader.AddFeatureClass(table, esriFeatureType.esriFTComplexEdge, uid, snap);
                    }
                    else {
                        // Add simple junction
                        UID uid = new UIDClass() {
                            Value = GUID_SIMPLEEDGE_CLSID
                        };
                        networkLoader.AddFeatureClass(table, esriFeatureType.esriFTSimpleEdge, uid, snap);
                    }
                }
            }

            // Check if network class has enabled field
            foreach (ZNetworkClass znc in zgn.NetworkClasses.Where(f =>!f.IsOrphanJunctionFeatureClass)) {
                string table = znc.Path.Table;
                string field = networkLoadedProps.DefaultEnabledField;

                switch (networkLoader.CheckEnabledDisabledField(table, field)) {
                    case esriNetworkLoaderFieldCheck.esriNLFCInvalidDomain:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - The enabled field has invalid domain", table, field), MessageType.Error);
                        return ResultType.Error;
                    case esriNetworkLoaderFieldCheck.esriNLFCInvalidType:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - The enabled field has invalid type", table, field), MessageType.Error);
                        return ResultType.Error;
                    case esriNetworkLoaderFieldCheck.esriNLFCNotFound:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - Adding enabled field", table, field), MessageType.Information);
                        networkLoader.PutEnabledDisabledFieldName(table, field);
                        break;
                    case esriNetworkLoaderFieldCheck.esriNLFCUnknownError:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - An unknown error was encountered", table, field), MessageType.Error);
                        return ResultType.Error;
                    case esriNetworkLoaderFieldCheck.esriNLFCValid:
                        GeometricNetworkViewModel.Default.AddMessage(string.Format("{0}:{1} - Enabled field already exists", table, field), MessageType.Warning);
                        break;
                }
            }

            // Import network weights
            foreach (ZNetWeight weight in zgn.Weights) {
                // Add weight
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding weight {0}", weight.Name), MessageType.Information);
                networkLoader.AddWeight(weight.Name, weight.WeightType.ToWeightType(), weight.BitGateSize);

                // Loop for each association
                foreach (ZNetWeightAssocation association in weight.NetWeightAssocations) {
                    // Add weight association
                    GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding weight association {0}:{1}", association.NetworkClass.Path.Table, association.Field.Name), MessageType.Information);
                    networkLoader.AddWeightAssociation(weight.Name, association.NetworkClass.Path.Table, association.Field.Name);
                }
            }

            // Load network
            try {
                networkLoader.LoadNetwork();
            }
            catch (Exception ex) {
                GeometricNetworkViewModel.Default.AddMessage(ex.Message, MessageType.Error);
                return ResultType.Error;
            }
            
            // Find new network
            IGeometricNetwork gn2 = w.FindGeometricNetwork(gxDialog.Name);
            if (gn2 == null) {
                GeometricNetworkViewModel.Default.AddMessage("Geometric network creation failed", MessageType.Error);
                return ResultType.Error;
            }

            // Update the name of the orphaned network class
            ZNetworkClass orphan = zgn.NetworkClasses.FirstOrDefault(n => n.IsOrphanJunctionFeatureClass);
            if (orphan != null) {
                orphan.Path.Table = string.Format("{0}_junctions", gxDialog.Name);
            }

            // Import junction rules
            foreach (ZJunctionConnectivityRule rule in zgn.JunctionRules) {
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding junction rule: {0}", rule.Id), MessageType.Information);
                IJunctionConnectivityRule2 r = new JunctionConnectivityRuleClass {
                    EdgeClassID = rule.Edge.Parent.ToVerfiedId(w).Value,
                    EdgeSubtypeCode = rule.Edge.Code,
                    JunctionClassID = rule.Junction.Parent.ToVerfiedId(w).Value,
                    JunctionSubtypeCode = rule.Junction.Code,
                    EdgeMinimumCardinality = rule.EdgeMinimum,
                    EdgeMaximumCardinality = rule.EdgeMaximum,
                    JunctionMinimumCardinality = rule.JunctionMinimum,
                    JunctionMaximumCardinality = rule.JunctionMaximum,
                    DefaultJunction = rule.IsDefault
                };
                gn2.AddRule(r);
            }

            // Import edge rules
            foreach (ZEdgeConnectivityRule rule in zgn.EdgeRules) {
                GeometricNetworkViewModel.Default.AddMessage(string.Format("Adding edge rule: {0}", rule.Id), MessageType.Information);
                IEdgeConnectivityRule r = new EdgeConnectivityRuleClass() {
                    FromEdgeClassID = rule.FromEdge.Parent.ToVerfiedId(w).Value,
                    FromEdgeSubtypeCode = rule.FromEdge.Code,
                    ToEdgeClassID = rule.ToEdge.Parent.ToVerfiedId(w).Value,
                    ToEdgeSubtypeCode = rule.ToEdge.Code
                };

                // Add junctions
                foreach (ZSubtype j in rule.Junctions) {
                    r.AddJunction(j.Parent.ToVerfiedId(w).Value, j.Code);
                }

                // Default junction?
                if (rule.DefaultJunction != null) {
                    r.DefaultJunctionClassID = rule.DefaultJunction.Parent.ToVerfiedId(w).Value;
                    r.DefaultJunctionSubtypeCode = rule.DefaultJunction.Code;
                }

                gn2.AddRule(r);
            }

            //
            GeometricNetworkViewModel.Default.AddMessage("Geometric network creation successful!", MessageType.Information);
            return ResultType.Successful;
        }    
        private static void TraverseVisualTree(DependencyObject reference, double offset) {
            var count = VisualTreeHelper.GetChildrenCount(reference);
            if (count == 0) { return; }
            for (int i = 0; i < count; i++) {
                var child = VisualTreeHelper.GetChild(reference, i);
                Grid grid = child as Grid;
                if (grid != null && grid.Tag != null && grid.Tag.ToString() == "groupheader") {
                    grid.Margin = new Thickness(offset, 0, 0, 0);
                }
                else {
                    GeometricNetworkWindow.TraverseVisualTree(child, offset);
                }
            }
        }
    }
}
