using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace RhPro.Oodd
{
    public partial class WaitIndicator : UserControl
    {
        #region Constructor
            public WaitIndicator()
            {
                InitializeComponent();

                if (!System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
                    LayoutRoot.Visibility = Visibility.Collapsed;
            }
        #endregion

        #region Public Functions
            public void Start()
            {
                LayoutRoot.Visibility = Visibility.Visible;
                IndicatorStoryboard.Begin();
            }

            public void Stop()
            {
                LayoutRoot.Visibility = Visibility.Collapsed;
                IndicatorStoryboard.Stop();
            }
        #endregion
    }
}
