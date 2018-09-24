using System;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using System.Text;
using System.Windows.Browser;
using System.Globalization;
using System.ServiceModel.DomainServices.Client;
using RhPro.Oodd.Utilities;
using RhPro.Oodd.Control.Node;
using RhPro.Oodd.Web.DataModel;
using RhPro.Oodd.Web.Services;
using System.Windows.Printing;
using System.Security;
using System.Windows.Media.Imaging;

namespace RhPro.Oodd
{
    public partial class MainPage : UserControl, IODPage
    {

        #region Constructors
        
        public MainPage()
        {
            InitializeComponent();

            this.setInitialParams();

            this.TextBoxRootDocket.Text = this.rootDocket.ToString();
            this.MouseMove += new MouseEventHandler(UserControl_MouseMove);
            this._dragDrop = new DragDropManager(this.CanvasTop);

            /* //imagen de fondo del organigrama, en la version 1 era un reloj en gris que no se pudo encontrar en los respaldos de XENNIT
            
            ImageBrush imageBrushBackground = new ImageBrush();
            imageBrushBackground.ImageSource = new BitmapImage(new Uri("images/pi.jpg", UriKind.Relative));
            imageBrushBackground.Stretch = Stretch.Fill;
            this.LiquidViewerMain.Background = imageBrushBackground;
            */

            bubblePopup.IsTimerEnabled = true;
            bubblePopup.TimeUntilClose = new TimeSpan(0, 0, 3);
            MessageBoxInfo.IsTimerEnabled = true;
            MessageBoxInfo.TimeUntilClose = new TimeSpan(0, 0, 10);

            this.ctlOrgChart.RootVisual = this;
            this.ctlOrgChart.CanvasPopupInfo = this.CanvasPopupInfo;
            this.ctlOrgChart.PointMousePosition = this.pointMousePosition;

            this.Loaded += OnLoaded;
        }

        #endregion

        #region Members

        private DragDropManager _dragDrop;
        private Point pointMousePosition;

        private XElement xeRoot;

        private Node root;
        private NodeCollection col = new NodeCollection();
        
        private enum operation { ReadOrg, ReadOrgFromPreviousEmp, ReadOrgFromNextEmp, SaveOrg };

        private int maxUndoLevel = 2;
        private int rootDocket = 0;
        private bool EnableTypeStructureTiles = false;
        private bool showCompleteName = false;
        private bool showTitles = false;
        private bool closeClickFlag = false;

        private Stack<RevertReferenceNode> undoCollection = new Stack<RevertReferenceNode>();
        private Stack<RevertReferenceNode> redoCollection = new Stack<RevertReferenceNode>();

        #endregion

        #region Properties

        private bool enableToolBar; 
        private bool enableSaveButton; 
        private bool enableRedoButton;
        private bool showWaiting;

        public bool EnableToolBar
        {
            get { 
                return enableToolBar; 
            }

            set {
                enableToolBar = value;
                this.ButtonNext.IsEnabled = value;
                this.ButtonPrevious.IsEnabled = value;
                this.ButtonPrint.IsEnabled = value;
                this.ButtonRedo.IsEnabled = false;
                this.ButtonSave.IsEnabled = false;
                this.ButtonUndo.IsEnabled = false;
                this.ButtonUpdate.IsEnabled = true;
                this.CanvasToolBar.UpdateLayout();
            }
        }

        public bool EnableRedoButton
        {
            get { return enableRedoButton; }
            set
            {
                enableRedoButton = value;
                this.ButtonRedo.IsEnabled = value;
            }
        }

        private bool ShowWaiting
        {
            get { return showWaiting; }
            set
            {
                showWaiting = value;
                if (value)
                {
                    this.ODWaits.Start();
                    this.ctlOrgChart.Visibility = Visibility.Collapsed;
                }
                else
                {
                    this.ODWaits.Stop();
                    this.ctlOrgChart.Visibility = Visibility.Visible;
                }
            }
        }
        
        #endregion

        #region Events

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            this.callToService(operation.ReadOrg, this.rootDocket, 5);
        }

        /// <summary>
        /// Maneja el evento al terminar la operacion de invocación del servicio y recupera el resultado 
        /// </summary>
        /// <param name="sender">Objeto que envía la solicitud</param>
        /// <param name="e">Datos del evento</param>
        private void OperationReadOrg_Completed(OrgChart orgChart)
        {
            if (orgChart != null)
            {
                if (orgChart.tree == null)
                {
                    if (orgChart.returnCode == 0)
                    {
                        this.ButtonUpdate_Click(null, null);
                    }
                    else
                    {
                        if (orgChart.returnCode < 0)
                        {
                            this.CanvasPopupWindow.Visibility = Visibility;

                            this.CanvasPopupWindow.Width = this.LiquidViewerMain.ActualWidth;
                            this.CanvasPopupWindow.Height = this.LiquidViewerMain.ActualHeight;

                            this.MessageBoxInfo.StartPosition = Liquid.DialogStartPosition.CenterParent;
                            this.MessageBoxInfo.ShowAsModal("\n" + orgChart.errorMessage, "RhPro X2 - Organigrama Dinámico");

                            this.MessageBoxInfo.Closed += new Liquid.DialogEventHandler(MessageBoxInfo_Closed);
                        }
                    }
                }
                else
                {
                    XElement tree = orgChart.tree;

                    this.loadTree(tree, null);

                    col.Clear();
                    col.AddNode(root);

                    this.drawChart();

                    this.TextBoxRootEmployee.Text = root.CompleteName;
                    this.TextBoxRootDocket.Text = root.Legajo;

                    this.EnableToolBar = true;
                    this.ButtonUpdate.IsEnabled = true;
                    
                    this.TextBoxRootDocket.IsEnabled = true;
                }
            }
            else
            {
                this.EnableToolBar = false;
                this.ButtonUpdate.IsEnabled = false;
            }

            this.ShowWaiting = false;
        }

        private void MessageBoxInfo_Closed(object sender, Liquid.DialogEventArgs e)
        {
            this.CanvasPopupWindow.Visibility = Visibility.Collapsed;
            this.CanvasPopupWindow.Width = 0;
            this.CanvasPopupWindow.Height = 0;

            this.EnableToolBar = true;
            this.ButtonUpdate.IsEnabled = true;
            
            this.TextBoxRootDocket.IsEnabled = true;
        }

        private void BubbleClose_Click(object sender, RoutedEventArgs e)
        {
            this.closeClickFlag = true;
            bubblePopup.Close();
        }

        [SecuritySafeCritical]
        private void ButtonPrint_Click(object sender, RoutedEventArgs e)
        {
            //this.StackPanelOrgChartContainer.PrintScreen();   
            //Lisandro Moro
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrintPage += OnPrintPage;
            printDoc.Print("RHPro - Organigrama Dinamico.");
        }

        void OnPrintPage(object sender, PrintPageEventArgs args){
            // Lisandro Moro - 08/10/2013
            // Combierto el stack a imagen para manipularle e imprimirla
            var bitmap = new WriteableBitmap(StackPanelOrgChartContainer, null);
            int MARGIN = 5;
            // Find the full size of the page
            Size pageSize =
              new Size(args.PrintableArea.Width
              + args.PageMargins.Left + args.PageMargins.Right,
              args.PrintableArea.Height
              + args.PageMargins.Top + args.PageMargins.Bottom);

            // Get additional margins to bring the total to MARGIN (= 96)
            Thickness additionalMargin = new Thickness{
                Left = Math.Max(0, MARGIN - args.PageMargins.Left),
                Top = Math.Max(0, MARGIN - args.PageMargins.Top),
                Right = Math.Max(0, MARGIN - args.PageMargins.Right),
                Bottom = Math.Max(0, MARGIN - args.PageMargins.Bottom)
            };

            // Find the area for display purposes
            Size displayArea =
              new Size(args.PrintableArea.Width
              - additionalMargin.Left - additionalMargin.Right,
              args.PrintableArea.Height
              - additionalMargin.Top - additionalMargin.Bottom);

            bool pageIsLandscape = displayArea.Width > displayArea.Height;
            bool imageIsLandscape = bitmap.PixelWidth > bitmap.PixelHeight;

            double displayAspectRatio = displayArea.Width / displayArea.Height;
            double imageAspectRatio = (double)bitmap.PixelWidth / bitmap.PixelHeight;

            double scaleX = Math.Min(1, imageAspectRatio / displayAspectRatio);
            double scaleY = Math.Min(1, displayAspectRatio / imageAspectRatio);

            // Calculate the transform matrix
            MatrixTransform transform = new MatrixTransform();

            if (pageIsLandscape == imageIsLandscape){
                // Pure scaling
                transform.Matrix = new Matrix(scaleX, 0, 0, scaleY, 0, 0);
            }else{
                // Scaling with rotation
                scaleX *= pageIsLandscape ? displayAspectRatio : 1 /
                  displayAspectRatio;
                scaleY *= pageIsLandscape ? displayAspectRatio : 1 /
                  displayAspectRatio;
                transform.Matrix = new Matrix(0, scaleX, -scaleY, 0, 0, 0);
            }

            Image image = new Image{
                Source = bitmap,
                Stretch = Stretch.Fill,
                Width = displayArea.Width,
                Height = displayArea.Height,
                RenderTransform = transform,
                RenderTransformOrigin = new Point(0.5, 0.5),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = additionalMargin,
            };

            Border border = new Border{
                Child = image,
            };

            args.PageVisual = border;
        } 

        private void loadTree(XElement tree, Node parent)
        {
            string firstName = tree.Attribute("nombre").Value.ToTitleCase();
            string lastName = tree.Attribute("apellido").Value.ToTitleCase();
            string space = string.Empty;

            string flName = string.Empty;

            if (this.showDock.ToLower().Equals("true")) flName = tree.Attribute("legajo").Value.ToString() + " | ";

            /*
            if (string.Concat(firstName, lastName).Length > 17)
                space = "\n";
            else
                space = " ";
            */

            if (!this.showCompleteName)
            {
                lastName = DataMan.cutWord(lastName);
                firstName = DataMan.cutWord(firstName);

                flName += firstName + " " + lastName;
            }
            else
            {
                flName += firstName;
            }

            string completeName = tree.Attribute("nombre").Value.ToTitleCase() + " " +
                tree.Attribute("apellido").Value.ToTitleCase();

            Node n = new Node()
            {
                ID = int.Parse(tree.Attribute("empCode").Value),
                Legajo = tree.Attribute("legajo").Value.ToString(),
                ParentNode = parent,
                Name = flName,
                LastName = lastName,
                CompleteName = completeName,
                IsExpanded = true,
                DragDrop = this._dragDrop,
                Company = tree.Attribute("empresa").Value.ToString().ToTitleCase(new List<char>() { '.', ' ' }),
                CompanyDesc = tree.Attribute("empresaDesc").Value.ToString().ToTitleCase(new List<char>() { '.', ' ' }), //Lisandro Moro
                RealCompany = tree.Attribute("empresa").Value.ToString().ToTitleCase(new List<char>() { '.', ' ' }),
                Division = tree.Attribute("sucursal").Value.ToString().ToTitleCase(),
                DivisionDesc = tree.Attribute("sucursalDesc").Value.ToString().ToTitleCase(), //Lisandro Moro
                Task = tree.Attribute("puesto").Value.ToString().ToTitleCase(),
                TaskDesc = tree.Attribute("puestoDesc").Value.ToString().ToTitleCase(), //Lisandro Moro
                Phone = tree.Attribute("interno").Value.ToString(),
                Email = tree.Attribute("mail").Value.ToTitleCase(),
                HeadShot = tree.Attribute("imageFileName").Value
            };

            if (this.showTitles)
            {
                if (!string.IsNullOrEmpty(n.Division))
                    //n.Division = "Sucursal: " + n.Division; //Lisandro Moro
                    n.Division = n.DivisionDesc + ": " + n.Division;
                if (!string.IsNullOrEmpty(n.Task))
                    //n.Task = "Puesto: " + n.Task; //Lisandro Moro
                    n.Task = n.TaskDesc + ": " + n.Task;
                if (!string.IsNullOrEmpty(n.Company))
                    //n.Company = "Empresa: " + n.Company; //Lisandro Moro
                    n.Company = n.CompanyDesc + ": " + n.Company;
            }

            if (parent == null)
                root = n;
            else
                parent.ChildNodes.AddNode(n);

            if (tree.HasElements)
                foreach (XElement xe in tree.Elements())
                    this.loadTree(xe, n);

        }

        private void ButtonUndo_Click(object sender, RoutedEventArgs e)
        {
            RevertReferenceNode rrn = this.undoCollection.Peek();
            if (rrn != null)
            {
                Node n = this.findNode(rrn.Id, root);
                Node newParentNode = this.findNode(rrn.IdOldParent, root);
                Node oldParentNode = this.findNode(rrn.IdNewParent, root);

                newParentNode.ChildNodes.Add(n);
                n.ParentNode = newParentNode;
                n.OldParentNode = oldParentNode;

                oldParentNode.ChildNodes.Remove(n);
                this.undoCollection.Pop();
                this.AddRedoNode(rrn);
                if (this.undoCollection.Count < 1) this.ButtonUndo.IsEnabled = false;
            }

            root.DragDrop.dropTargets.Clear();
            ctlOrgChart.FillChart();
        }

        private void ButtonRedo_Click(object sender, RoutedEventArgs e)
        {
            RevertReferenceNode rrn = this.redoCollection.Peek();
            if (rrn != null)
            {
                Node n = this.findNode(rrn.Id, root);
                Node oldParentNode = this.findNode(rrn.IdOldParent, root);
                Node newParentNode = this.findNode(rrn.IdNewParent, root);

                newParentNode.ChildNodes.Add(n);
                n.ParentNode = newParentNode;
                n.OldParentNode = oldParentNode;

                oldParentNode.ChildNodes.Remove(n);
                this.redoCollection.Pop();
                this.AddUndoNode(rrn);
                if (this.redoCollection.Count < 1) this.ButtonRedo.IsEnabled = false;
            }

            root.DragDrop.dropTargets.Clear();
            ctlOrgChart.FillChart();

        }

        private void UserControl_MouseMove(object sender, MouseEventArgs e)
        {
            Point p = e.GetPosition(this);
            
            this.TextBlockPositionX.Text = p.X.ToString();
            this.TextBlockPositionY.Text = p.Y.ToString();
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            this.xeRoot = new XElement("nodes");
            this.BuildOrgChart(root);
            this.callToService(operation.SaveOrg, 0, 5);

            this.undoCollection.Clear();
            this.redoCollection.Clear();
        }

        private void BuildOrgChart(Node node)
        {
            if (node == null) node = this.root;
            if (node.HasChildren)
            {
                foreach (Node n in node.ChildNodes)
                {
                    if (n.OldParentNode != null)
                    {
                        XElement xe = new XElement("node");
                        xe.SetAttributeValue("codEmp", n.ID.ToString());
                        xe.SetAttributeValue("oldCodParent", n.OldParentNode.ID.ToString());
                        xe.SetAttributeValue("newCodParent", n.ParentNode.ID.ToString());

                        xeRoot.Add(xe);
                    }
                    this.BuildOrgChart(n);

                }
            }
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            this.callToService(operation.ReadOrg, int.Parse(this.TextBoxRootDocket.Text),
            int.Parse(this.NumericUpDownNodeLevel.Value.ToString()));
        }

        private void TextBoxRootDocket_KeyDown(object sender, KeyEventArgs e)
        {
            DataMan.allowNumbers(e);
            if (e.PlatformKeyCode == 13 || e.PlatformKeyCode == 9)
            {
                this.TextBoxRootEmployee.Text = string.Empty;
                this.ButtonUpdate_Click(sender, e);
            }
        }

        private void NumericUpDownNodeLevel_KeyDown(object sender, KeyEventArgs e)
        {
            DataMan.allowNumbers(e);

            if (e.PlatformKeyCode == 13)
                this.ButtonUpdate_Click(sender, e);
        }

        private void ButtonPrevious_Click(object sender, RoutedEventArgs e)
        {
            this.callToService(operation.ReadOrgFromPreviousEmp, int.Parse(this.TextBoxRootDocket.Text),
            int.Parse(this.NumericUpDownNodeLevel.Value.ToString()));
        }

        private void ButtonNext_Click(object sender, RoutedEventArgs e)
        {
            this.callToService(operation.ReadOrgFromNextEmp, int.Parse(this.TextBoxRootDocket.Text),
            int.Parse(this.NumericUpDownNodeLevel.Value.ToString()));
        }

        private void TextBoxRootDocket_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            tb.SelectAll();
        }

        private void LiquidViewerMain_LayoutUpdated(object sender, EventArgs e)
        {
            this.TextBlockZoom.Text = this.LiquidViewerMain.Zoom.ToString();
            if (!closeClickFlag)
                this.BubbleClose_Click(null, null);
        }

        private void ImageInfo_ImageFailed(object sender, ExceptionRoutedEventArgs e)
        {
            ImageInfo.Source = null;
        }

        #endregion

        #region Methods

        private void drawChart()
        {
            this.root.DragDrop.dropTargets.Clear();

            this.ctlOrgChart.Nodes = col;
            this.ctlOrgChart.FillChart();

            this.LiquidViewerMain.ScrollIntoPosition(this.root);
        }

        /// <summary>
        /// Invoca al servicio
        /// </summary>
        /// <param name="operation">Operación que se va a invocar</param>
        /// <param name="legajo">Nro de Legajo raíz</param>
        /// <param name="maxLevel">Cantidad de niveles permitidos</param>
        private void callToService(operation operation, int legajo, int maxLevel)
        {
            this.EnableToolBar = false;
            this.ButtonUpdate.IsEnabled = false;
            //this.NumericUpDownNodeLevel.IsEnabled = false;
            this.TextBoxRootDocket.IsEnabled = false;
            try
            {
                //OrgDomainContext proxyService = new OrgDomainContext();

                RhPro.Oodd.Web.Services.OdDomainContext proxyService = new OdDomainContext();
                
                InvokeOperation<OrgChart> operationRead = null;

                switch (operation)
                {
                    case operation.ReadOrg:
                        operationRead = proxyService.ReadOrg(legajo, maxLevel);
                        break;
                    case operation.ReadOrgFromPreviousEmp:
                        operationRead = proxyService.ReadOrgFromPreviousEmp(legajo, maxLevel);
                        break;
                    case operation.ReadOrgFromNextEmp:
                        operationRead = proxyService.ReadOrgFromNextEmp(legajo, maxLevel);
                        break;
                    case operation.SaveOrg:
                        operationRead = proxyService.SaveOrg(this.xeRoot);
                        break;
                }

                this.ShowWaiting = true;
                //operationRead.Completed += new EventHandler(OperationReadOrg_Completed);
                operationRead.Completed += (s, args) =>
                    {
                        if (operationRead.HasError) {
                            operationRead.MarkErrorAsHandled();
                        }
                        else {
                            OperationReadOrg_Completed(operationRead.Value);
                        }
                    };
                
                
                
                
                
                

            }
            catch (Exception ex)
            {
                //Output.Text = "Error en la invocacion";
            }
        }

        /// <summary>
        /// Devuelve un grupo de transformaciones para adaptar el organigrama a una hoja de impresión
        /// </summary>
        /// <returns>Grupo de transformaciones</returns>
        private TransformGroup imageTransform()
        {
            ScaleTransform st = new ScaleTransform();
            double scale = (double)500 / (double)Math.Max(
                this.StackPanelOrgChartContainer.ActualHeight, this.StackPanelOrgChartContainer.ActualWidth);
            st.CenterX = 0;
            st.CenterY = 0;
            st.ScaleX = scale;
            st.ScaleY = scale;

            RotateTransform rt = new RotateTransform();
            rt.Angle = 270;
            rt.CenterX = st.ScaleX / 2;
            rt.CenterY = st.ScaleY / 2;

            TranslateTransform tt = new TranslateTransform();
            //tt.Y = 500;
            tt.Y = this.StackPanelOrgChartContainer.ActualWidth;


            TransformGroup tg = new TransformGroup();
            //tg.Children.Add(st);
            tg.Children.Add(rt);
            tg.Children.Add(tt);




            return tg;
        }

        private Node findNode(int Id, Node node)
        {
            if (node.ID != Id)
            {
                foreach (Node n in node.ChildNodes)
                {
                    if (n.ID == Id)
                        return n;
                    Node n2 = this.findNode(Id, n);
                    if (n2 != null)
                        return n2;
                }
            }
            else return node;

            return null;
        }

        private void setInitialParams()
        {

            foreach (KeyValuePair<string, string> s in Application.Current.Host.InitParams)
            {
                if (s.Key.Equals("showDock"))
                    this.showDock = s.Value;

                if (s.Key.Equals("imagePath"))
                    this.imagePath = s.Value;

                //if (s.Key.Equals("maxUndoLevel"))
                //this.maxUndoLevel = int.Parse(s.Value)+1;

                if (s.Key.Equals("rootDocket"))
                    this.rootDocket = int.Parse(s.Value);

                if (s.Key.Equals("showCompleteName"))
                    this.showCompleteName = bool.Parse(s.Value);

                if (s.Key.Equals("showTitles"))
                    this.showTitles = bool.Parse(s.Value);

            }

            if (string.IsNullOrEmpty(this.imagePath)) this.imagePath = "../FOTOS/";
            if (this.maxUndoLevel > 9) this.maxUndoLevel = 9;
        }

        float DeterminePercentageForResize(int height, int width)
        {
            int highestValue;
            if (height > width)
                highestValue = height;
            else
                highestValue = width;

            float percent = 100 / (float)highestValue;

            if (percent > 1 && percent != 0)
                throw new Exception("Percent cannot be greater than 1 or equal to zero");
            else
                return percent;
        }

        #endregion

        #region IPage Members

        private bool enableUndoButton = false;

        bool IODPage.EnableSaveButton
        {
            get { return enableSaveButton; }
            set
            {
                enableSaveButton = value;
                this.ButtonSave.IsEnabled = value;
            }
        }

        bool IODPage.EnableUndoButton
        {
            get
            {
                return enableUndoButton;
            }
            set
            {
                enableRedoButton = value;
                this.ButtonUndo.IsEnabled = value;
            }
        }

        void AddUndoNode(RevertReferenceNode undoNode)
        {
            if (this.undoCollection.Count > this.maxUndoLevel)
                this.undoCollection.Pop();

            this.undoCollection.Push(undoNode);

            this.ButtonUndo.IsEnabled = true;
        }

        void AddRedoNode(RevertReferenceNode redoNode)
        {
            if (this.redoCollection.Count > this.maxUndoLevel)
                this.redoCollection.Pop();

            this.redoCollection.Push(redoNode);

            this.ButtonRedo.IsEnabled = true;
        }

        #endregion

        #region IODPage Members

        private string imagePath = string.Empty;
        private string showDock = string.Empty;

        void IODPage.AddUndoNode(RevertReferenceNode undoNode)
        {
            this.AddUndoNode(undoNode);
        }

        void IODPage.AddRedoNode(RevertReferenceNode redoNode)
        {
            this.AddRedoNode(redoNode);
        }

        string IODPage.ShowDock
        {
            get { return showDock; }
        }

        int IODPage.MaxUndoLevel
        {
            get { return this.maxUndoLevel; }
        }

        bool IODPage.ShowCompleteName
        {
            get { return this.showCompleteName; }
        }

        string IODPage.ImagePath
        {
            get { return imagePath; }
        }

        #endregion

    }
}
