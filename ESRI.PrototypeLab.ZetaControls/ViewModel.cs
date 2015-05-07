/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Windows.Threading;

namespace ESRI.PrototypeLab.ZetaControls {
    public abstract class ViewModel : DependencyObject {
        //
        // CONSTRUCTOR
        //
        public ViewModel() {
            this.Messages = new ObservableCollection<MessageItem>();
        }
        //
        // STATIC
        //
        public static readonly DependencyProperty IsDirtyProperty = DependencyProperty.Register(
            "IsDirty",
            typeof(bool),
            typeof(ViewModel),
            new PropertyMetadata(false)
        );
        public static readonly DependencyProperty DatasetProperty = DependencyProperty.Register(
            "Dataset",
            typeof(ZDataset),
            typeof(ViewModel),
            new PropertyMetadata(null)
        );
        public static readonly DependencyProperty DocumentProperty = DependencyProperty.Register(
            "Document",
            typeof(string),
            typeof(ViewModel),
            new PropertyMetadata(null)
        );
        //
        // PROPERTIES
        //
        public ObservableCollection<MessageItem> Messages { get; private set; }
        public string WindowTitle { get; protected set; }
        public ZDataset Dataset {
            get { return (ZDataset)this.GetValue(ViewModel.DatasetProperty); }
            set { this.SetValue(ViewModel.DatasetProperty, value); }
        }
        public bool IsDirty {
            get { return (bool)this.GetValue(ViewModel.IsDirtyProperty); }
            set { this.SetValue(ViewModel.IsDirtyProperty, value); }
        }
        public string Document {
            get { return (string)this.GetValue(ViewModel.DocumentProperty); }
            set { this.SetValue(ViewModel.DocumentProperty, value); }
        }
        //
        // METHODS
        //
        public void AddMessage(string message, MessageType type) {
            this.Dispatcher.Invoke(
               DispatcherPriority.Normal,
               new Action(
                   delegate() {
                       switch (type) {
                           case MessageType.Information:
                               this.Messages.Add(
                                   new InformationItem() {
                                       Message = message
                                   }
                               );
                               break;
                           case MessageType.Warning:
                               this.Messages.Add(
                                   new WarningItem() {
                                       Message = message
                                   }
                               );
                               break;
                           case MessageType.Error:
                               this.Messages.Add(
                                   new ErrorItem() {
                                       Message = message
                                   }
                               );
                               break;
                       }
                   }
               )
           );
        }
        public virtual void Clear() {
            this.Dataset = null;
            this.IsDirty = false;
            this.Document = null;
        }
        public void MakeDirty() {
            this.IsDirty = true;
        }
        public void Save() {
            this.Save(this.Document);
        }
        public void Save(string document) {
            FileStream fs = new FileStream(document, FileMode.Create);
            BinaryFormatter formatter = new BinaryFormatter();
            try {
                ZDataset d = this.Dataset;
                formatter.Serialize(fs, d);
                this.Document = document;
                this.IsDirty = false;
            }
            catch (SerializationException e) {
                string message = e.Message;
                if (string.IsNullOrWhiteSpace(message)) {
                    message = "Error saving document";
                }
                MessageBoxResult r = MessageBox.Show(
                    message,
                    GeometricNetworkViewModel.Default.WindowTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK
                );
            }
            finally {
                fs.Close();
            }
        }
        public virtual void Load(string document) {
            FileStream fs = new FileStream(document, FileMode.Open);
            try {
                BinaryFormatter formatter = new BinaryFormatter();
                this.Dataset = (ZDataset)formatter.Deserialize(fs);
                this.Document = document;
            }
            catch (SerializationException e) {
                string message = e.Message;
                if (string.IsNullOrWhiteSpace(message)) {
                    message = "Error opening document";
                }
                MessageBoxResult r = MessageBox.Show(
                    message,
                    GeometricNetworkViewModel.Default.WindowTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK
                );
            }
            finally {
                fs.Close();
            }
        }
    }
}
