﻿#pragma checksum "D:\Documents\Visual Studio 2010\Projects\RhPro.Oodd\RhPro.Oodd\WaitIndicator.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "878F77B39C87590133BC01C34D734EB3"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.17929
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Resources;
using System.Windows.Shapes;
using System.Windows.Threading;


namespace RhPro.Oodd {
    
    
    public partial class WaitIndicator : System.Windows.Controls.UserControl {
        
        internal System.Windows.Media.Animation.Storyboard IndicatorStoryboard;
        
        internal System.Windows.Controls.Canvas LayoutRoot;
        
        internal System.Windows.Shapes.Ellipse Ellipse1;
        
        internal System.Windows.Shapes.Ellipse Ellipse2;
        
        internal System.Windows.Shapes.Ellipse Ellipse3;
        
        internal System.Windows.Shapes.Ellipse Ellipse4;
        
        internal System.Windows.Shapes.Ellipse Ellipse5;
        
        internal System.Windows.Shapes.Ellipse Ellipse6;
        
        internal System.Windows.Shapes.Ellipse Ellipse7;
        
        internal System.Windows.Shapes.Ellipse Ellipse8;
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Windows.Application.LoadComponent(this, new System.Uri("/RhPro.Oodd;component/WaitIndicator.xaml", System.UriKind.Relative));
            this.IndicatorStoryboard = ((System.Windows.Media.Animation.Storyboard)(this.FindName("IndicatorStoryboard")));
            this.LayoutRoot = ((System.Windows.Controls.Canvas)(this.FindName("LayoutRoot")));
            this.Ellipse1 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse1")));
            this.Ellipse2 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse2")));
            this.Ellipse3 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse3")));
            this.Ellipse4 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse4")));
            this.Ellipse5 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse5")));
            this.Ellipse6 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse6")));
            this.Ellipse7 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse7")));
            this.Ellipse8 = ((System.Windows.Shapes.Ellipse)(this.FindName("Ellipse8")));
        }
    }
}

