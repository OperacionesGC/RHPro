using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace RhPro.Oodd.Control.Node
{
    /// <summary>
    /// Extiende StackPanel a fin de que conozca a su propio nodo
    /// </summary>
    public class StackPanelNode : StackPanel
    {
        public Node NodeOwner { get; set; }
    }
}
