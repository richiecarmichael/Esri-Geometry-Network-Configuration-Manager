/* ----------------------------------------------- 
 * Copyright © 2013 Esri Inc. All Rights Reserved. 
 * ----------------------------------------------- */

using System;

namespace ESRI.PrototypeLab.ZetaControls {
    public class MessageEventArgs : EventArgs {
        public MessageType MessageType { get; set; }
        public string Message { get; set; }
    }
}
