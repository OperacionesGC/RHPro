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
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using System.ComponentModel;

using RhPro.Oodd.Utilities;

namespace RhPro.Oodd.Control.Node
{
    public partial class NodeContent : UserControl
    {
        

        /*
         * verde 163 138 30
         * amarillo 201 125 24
         * azul 53 90 150
         * rosa 201 95 138
         * otro 167 119 252
         */

        /*
         * verde 163 138 30 A38A1E
         * amarillo 201 125 24 C97D18
         * azul 36 79 150 244F96
         * rosa 201 71 104 C94768
         * otro 136 85 252 8855FC
         */
        
        #region Private Members
        private Node _currentNode;

        private Color _rootElement = Color.FromArgb(254, 91, 237, 57); //
        private Color _leafElement = Color.FromArgb(254, 136, 173, 249); //
        private Color _defaultElement = Color.FromArgb(254, 251, 251, 118); //
        public Color _selectedElement = Color.FromArgb(154, 254, 1, 244); // public para poder pintarlo desde afuera
        private Color _defaultBorder = Color.FromArgb(102, 0, 0, 0); //
        private Color _hoverBorder = Colors.Red; //

        #endregion

        public Border BrdLowerRectangle;

        #region Constructors
        public NodeContent(Node current)
        {
            InitializeComponent();
            
            this.Width = current.Width;
            this.Height = current.Height;
            
            if (current != null)
            {
                string division = current.Division;
                string task = current.Task;
                string company = current.Company;

                bool reserveSpaces = true;
                bool showCompleteName = false;

                IDictionary<string, string> initParams = Application.Current.Host.InitParams;

                if (initParams.Keys.Contains("reserveSpaces")) reserveSpaces = bool.Parse(initParams["reserveSpaces"].ToLower());
                if (initParams.Keys.Contains("showCompleteName")) showCompleteName = bool.Parse(initParams["showCompleteName"].ToLower());

                if (reserveSpaces)
                {
                    if (division.Equals(string.Empty)) division = " ";
                    if (task.Equals(string.Empty)) task = " ";
                    if (company.Equals(string.Empty)) company = " ";
                }

                if (!showCompleteName)
                {
                    this.brdLowerRectangle.MaxWidth = 150;
                    //this.brdLowerRectangle.Margin = new Thickness(5, 0, 5, 0);
                }
                else
                {
                    //this.lblName.Height = 30;
                    //this.lblName.Width = 95;
                }

                this.lblName.Content = current.Legajo + " - " + current.Name;

                if (!showCompleteName)
                {
                    this.lblLastName.Content = "";
                    this.lblLastName.Visibility = Visibility.Collapsed;
                }
                else
                {
                    this.lblLastName.Visibility = Visibility.Visible;
                    this.lblLastName.Content = current.LastName;
                }
                
                this.lblPosition0.Text = company;
                this.lblPosition1.Text = division;
                this.lblPosition2.Text = task;

                _currentNode = current;
            }

            
        }
        #endregion

        #region Event Handling
        private void OnMouseEnter(object sender, MouseEventArgs e)
        {
            brdHover.BorderBrush = new SolidColorBrush(_hoverBorder);
        }

        private void OnMouseLeave(object sender, MouseEventArgs e)
        {
            brdHover.BorderBrush = new SolidColorBrush(_defaultBorder);
        }

        private void OnStateButtonClicked(object sender, RoutedEventArgs e)
        {
            _currentNode.IsExpanded = !_currentNode.IsExpanded;
            SetButtonVisibility();
        }
        #endregion

        #region Public Methods
        public void Initiate()
        {
            if (_currentNode.ParentNode == null)
                brdMain.Background = this.createGradientColor(_rootElement); // new SolidColorBrush(_rootElement);
            else if (!_currentNode.HasChildren)
                brdMain.Background = this.createGradientColor(_leafElement);//new SolidColorBrush(_leafElement);
            else
                brdMain.Background = this.createGradientColor(_defaultElement); //new SolidColorBrush(_defaultElement);


            brdHover.BorderBrush = new SolidColorBrush(_defaultBorder);
            
            SetButtonVisibility();
        }

        private Brush createGradientColor(Color color)
        {
            LinearGradientBrush r = new LinearGradientBrush();

            r.StartPoint = new Point(0, 0);
            r.EndPoint = new Point(0, 1);

            r.GradientStops.Add(new GradientStop() { Color = Colors.White, Offset = 0.0 });
            r.GradientStops.Add(new GradientStop() { Color = color, Offset = 1.0 });

            return r;
        }

        public void Reset()
        {
            Initiate();
        }

        public void Select()
        {
            brdMain.Background = new SolidColorBrush(_selectedElement);
        }
        #endregion

        #region Private Methods
        private void SetButtonVisibility()
        {
            if (_currentNode.HasChildren)
            {
                if (_currentNode.IsExpanded)
                {
                    btnMaximize.Opacity = 0;
                    btnMinimize.Opacity = 1;
                    btnMaximize.Visibility = Visibility.Collapsed;
                    btnMinimize.Visibility = Visibility.Visible;
                    lblName.Padding =  this.NewButtonTopMeasure(0);
                }
                else
                {
                    btnMaximize.Opacity = 1;
                    btnMinimize.Opacity = 0;
                    btnMinimize.Visibility = Visibility.Collapsed;
                    btnMaximize.Visibility = Visibility.Visible;
                    lblName.Padding = this.NewButtonTopMeasure(0);
                }
            }
            else
            {
                btnMaximize.Opacity = 0;
                btnMinimize.Opacity = 0;
                btnMaximize.Visibility = Visibility.Collapsed;
                btnMinimize.Visibility = Visibility.Collapsed;
                lblName.Margin = new Thickness(0, 15, 0, 0);
            }
        }

        private Thickness NewButtonTopMeasure(double topValue)
        {
            return new Thickness(this.lblName.Padding.Left, topValue, 
                this.lblName.Padding.Right, this.lblName.Padding.Bottom);

        }
        #endregion

        Image imageInfo;
        private void lblName_Click(object sender, RoutedEventArgs e)
        {
            Image imageInfoReflected;
            Canvas canvasPopupInfo = this._currentNode.OrgChart.CanvasPopupInfo;
            canvasPopupInfo.Visibility = Visibility;
            //Point pointMousePosition = this._currentNode.OrgChart.PointMousePosition;

            Liquid.Bubble bubblePopup = canvasPopupInfo.Children[0] as Liquid.Bubble;
            
            if (bubblePopup != null)
            {
                TextBlock x = canvasPopupInfo.Children[1] as TextBlock;
                TextBlock y = canvasPopupInfo.Children[2] as TextBlock;
                bubblePopup.VerticalOffset = double.Parse(y.Text)+8; //pointMousePosition.Y;
                bubblePopup.HorizontalOffset = double.Parse(x.Text)+8;// pointMousePosition.X;

                Grid gridInfo = ((bubblePopup.Content as StackPanel).Children[0] as Grid);

                //imageInfo = ((bubblePopup.Content as StackPanel).Children[0] as Border).Child as Image;
                //imageInfo = (bubblePopup.Content as StackPanel).Children[0] as Image;
                //Canvas canvasImage = gridInfo.Children[0] as Canvas;
                imageInfo = gridInfo.Children[0] as Image;
                imageInfoReflected = gridInfo.Children[1] as Image;

                if (imageInfo != null)
                {
                    string stringHeadShot = this._currentNode.HeadShot;

                    if (stringHeadShot.Equals(string.Empty))
                    {
                        stringHeadShot = "ImageDefaultUser.jpg";
                    }

                    string imagePath = this._currentNode.OrgChart.RootVisual.ImagePath;

                    if (!imagePath.Substring(imagePath.Length - 1, 1).Equals("/"))
                    {
                        imagePath += "/";
                    }

                    Uri uri = null;

                    if (isSupportedFormat(stringHeadShot))
                    {
                        uri = new Uri(Application.Current.Host.Source, imagePath + stringHeadShot);
                        BitmapImage bmp = new BitmapImage();
                        bmp.UriSource = uri;
                        imageInfo.Source = bmp;
                    }
                    else
                    {
                        string urlPrefix = "http://";
                        urlPrefix = urlPrefix + Application.Current.Host.Source.Host;
                        urlPrefix = urlPrefix + ":" + Application.Current.Host.Source.Port;
                        string absolutePath = Application.Current.Host.Source.AbsolutePath;
                        absolutePath = absolutePath.Remove(0, 1);
                        string[] tokens = absolutePath.Split('/');
                        int i = 0;
                        if (tokens.Length > 2)
                        {
                            while (i < tokens.Length - 2)
                            {
                                urlPrefix = urlPrefix + "/" + tokens[i];
                                i++;
                            }
                        }

                        uri = new Uri(urlPrefix + "/ImageHttpHandler.ashx?imageName=" + stringHeadShot);

                        imageInfo.Source = new BitmapImage(uri);
                    }

                    
                    imageInfo.Stretch = Stretch.Uniform;

                    
                }

                TextBlock textBlockName = gridInfo.Children[2] as TextBlock;
                TextBlock textBlockPhoneTitle = gridInfo.Children[3] as TextBlock;
                TextBlock textBlockPhone = gridInfo.Children[4] as TextBlock;
                TextBlock textBlockEmailTitle = gridInfo.Children[5] as TextBlock;
                TextBlock textBlockEmail = gridInfo.Children[6] as TextBlock;
                TextBlock textBlockEmpresaTitle = gridInfo.Children[7] as TextBlock;
                TextBlock textBlockEmpresa = gridInfo.Children[8] as TextBlock;

                if (textBlockName != null)
                    textBlockName.Text = this._currentNode.CompleteName;
                
                if (!this._currentNode.Phone.Equals(string.Empty)){
                    
                    textBlockPhoneTitle.Visibility = Visibility.Visible;
                    if (textBlockPhone != null) textBlockPhone.Text = this._currentNode.Phone;
                }else{
                    textBlockPhoneTitle.Visibility = Visibility.Collapsed;
                    textBlockPhone.Text = string.Empty;
                }

                if (!this._currentNode.Email.Equals(string.Empty)){
                    
                    textBlockEmailTitle.Visibility = Visibility.Visible;
                    if (textBlockEmail != null) textBlockEmail.Text = this._currentNode.Email;
                }else{
                    textBlockEmailTitle.Visibility = Visibility.Collapsed;
                    textBlockEmail.Text = string.Empty;
                }

                if(!this._currentNode.RealCompany.Equals(string.Empty)){
                    
                    textBlockEmpresaTitle.Visibility = Visibility.Visible;
                    if (textBlockEmpresa != null) textBlockEmpresa.Text = this._currentNode.RealCompany;
                }else{
                    textBlockEmpresaTitle.Visibility = Visibility.Collapsed;
                    textBlockEmpresa.Text = string.Empty;
                }

                //imageInfoReflected.Margin = new Thickness(0, imageInfo.DesiredSize.Height, 0, 0);
                bubblePopup.Show();
            }
        }

        protected bool isSupportedFormat(string filename)
        {
            string[] tokens = filename.Split('.');
            if (tokens.Length == 2)
            {
                string format = tokens[1].ToLower();
                if (format.Equals("jpg") || format.Equals("png"))
                {
                    return true;
                }
                else return false;
            }
            return false;
        }

        void bitmapImage_DownloadProgress(object sender, DownloadProgressEventArgs e)
        {
            if (e.Progress == 100)
            {
                Dispatcher.BeginInvoke(delegate()
                {
                    double height = imageInfo.ActualHeight;
                    double width =  imageInfo.ActualWidth;
                });
            }
        }


        
    }
}
