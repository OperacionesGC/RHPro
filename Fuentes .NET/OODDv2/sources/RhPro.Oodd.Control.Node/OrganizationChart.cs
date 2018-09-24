using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Liquid;
using RhPro.Oodd.Utilities;

namespace RhPro.Oodd.Control.Node
{
    public class OrganizationChart : Canvas
    {
        #region Private Members
        private double _hSpace = 5;
        private double _vSpace = 30;
        public Node _selectedNode; //se debe conocer para poder deseleccionarlo
        private Node _rootNode;
        private NodeCollection _nodes;
        public Canvas CanvasPopupInfo;
        public Point PointMousePosition;
        #endregion

        #region Public Properties
            public NodeCollection Nodes
            {
                set { _nodes = value; }
            }
            public string ImageLocation { get; set; }
            public IODPage RootVisual { get; set; }
        #endregion

        #region Constructors
        public OrganizationChart()
        {
            _nodes = new NodeCollection();
            Loaded += OnLoaded;
        }

        public OrganizationChart(NodeCollection nodes)
        {
            _nodes = nodes;
            Loaded += OnLoaded;
        }
        #endregion

        #region Event Handling
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_nodes != null && _nodes.Count > 0)
                FillChart();
        }

        public bool boolSelectionNodeFlag = false; //evita que se pase dos veces sobre el mismo click

        private void OnNodeSelectionChanged(Node selectedNode)
        {
            /*
            if (_selectedNode != null)
            {
                _selectedNode.Content.Reset();

                if (!_selectedNode.Equals(selectedNode))
                {
                    _selectedNode = selectedNode;
                    _selectedNode.Content.Select();
                }else
                    _selectedNode = null;
            }else{
                _selectedNode = selectedNode;
                _selectedNode.Content.Select();    
            }
            */
            /*
            if (_selectedNode != null)
                _selectedNode.Content.Reset();

            _selectedNode = selectedNode;
            _selectedNode.Content.Select();    
            */
        }


        private void OnNodeStateChanged(Node currentNode, bool isExpanded)
        {
            DrawChart();
        }
        #endregion

        #region Public Methods

        public void FillChart()
        {
            if (_nodes.Count == 1)
            {
                _rootNode = _nodes[0];

                InitiateNode(_rootNode);
                DrawChart();
            }
        }

        public void DeselectNode(Node selectNode)
        {
            this.OnNodeSelectionChanged(selectNode);
        }

        /// <summary>
        /// Dibuja el organization chart desde el nodo indicado
        /// </summary>
        /// <param name="fromNode">Nodo desde el cual se comienza a re-dibujar el organigrama</param>
        public void FillChart(Node fromNode)
        {
            InitiateNode(fromNode);
            DrawChart(fromNode);
        }

        public void ChangeNodeState(Node currentNode, bool isExpanded)
        {
            currentNode.IsExpanded = isExpanded;
        }
        #endregion

        #region Private Methods
        private void InitiateNode(Node currentNode)
        {
            currentNode.NodeSelectionChanged += OnNodeSelectionChanged;
            currentNode.NodeStateChanged += OnNodeStateChanged;
            currentNode.OrgChart = this;
            
            currentNode.Content = new NodeContent(currentNode);
            currentNode.Content.StackPanelDrag.NodeOwner = currentNode;
            currentNode.Content.StackPanelDrop.NodeOwner = currentNode;

            currentNode.Content.Initiate();

            if (currentNode.HasChildren)
                foreach (Node node in currentNode.ChildNodes)
                    InitiateNode(node);
        }

        private void DrawChart()
        {
            this.Children.Clear();

            Measure();

            double[] nodesWOC = new double[1];

            PlaceNode(_rootNode, 1, nodesWOC);
            DrawConnectionLines(_rootNode);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fromNode"></param>
        private void DrawChart(Node fromNode)
        {
            //this.Children.Clear();

            Measure();

            double[] nodesWOC = new double[1];

            PlaceNode(fromNode, 1, nodesWOC);
            DrawConnectionLines(fromNode);
        }

        private void Measure()
        {
            if (_rootNode != null)
            {
                int levels = CountLevels(_rootNode, 1);
                int finalNodes = CountNodesWithoutChildren(_rootNode);

                double nodeWidth = _rootNode.Width;
                double compWidth = finalNodes * nodeWidth + (finalNodes + 1) * _hSpace;

                double nodeHeight = _rootNode.Height;
                double compHeight = levels * nodeHeight + (levels + 1) * _vSpace;

                this.Width = compWidth;
                this.Height = compHeight;
            }
        }

        private int CountNodesWithoutChildren(Node node)
        {
            if (node != null && node.IsExpanded)
            {
                if (!node.HasChildren)
                {
                    return 1;
                }
                else
                {
                    int count = 0;
                    for (int i = 0; i <= node.ChildNodes.Count - 1; i++)
                    {
                        Node child = node.ChildNodes[i];
                        count += CountNodesWithoutChildren(child);
                    }

                    return count;
                }
            }

            return 1;
        }

        private int CountLevels(Node node, int level)
        {
            if (node != null && node.IsExpanded)
            {
                if (!node.HasChildren)
                {
                    return level;
                }
                else
                {
                    int maxLevel = level;
                    for (int i = 0; i <= node.ChildNodes.Count - 1; i++)
                    {
                        Node child = node.ChildNodes[i];
                        maxLevel = Math.Max(maxLevel, CountLevels(child, (level + 1)));
                    }

                    return maxLevel;
                }
            }

            return level;
        }

        private Point PlaceNode(Node node, int level, double[] nodesWOC)
        {
            double nodeHeight = node.Height;
            double nodeWidth = node.Width;

            if (!node.HasChildren || !node.IsExpanded)
            {
                nodesWOC[0] = (double)nodesWOC[0] + 1;

                double left = (nodeWidth + _hSpace) * ((double)nodesWOC[0] - 1) + _hSpace;
                double top = (nodeHeight + _vSpace) * (level - 1) + _vSpace;

                node.Position = new Point(left, top);
                
                Canvas.SetLeft(node, left);
                Canvas.SetTop(node, top);

                //si el nodo está ya contenido antes debo eliminarlo, para luego si agregarlo
                if (this.Children.Contains(node)) this.Children.Remove(node);
                this.Children.Add(node);
                return node.Position;
            }
            else
            {
                double posX = 0, posY = 0;
                Point first = new Point();
                Point last = new Point();
                Point temp = new Point();

                for (int i = 0; i <= (node.ChildNodes.Count - 1); i++)
                {
                    Node n = node.ChildNodes[i];
                    temp = PlaceNode(n, level + 1, nodesWOC);

                    if (i == 0)
                        first = temp;

                    if (i == (node.ChildNodes.Count - 1))
                        last = temp;
                }

                posX = (first.X + last.X) / 2;
                posY = (nodeHeight + _vSpace) * (level - 1) + _vSpace;

                //if (node.Equals(_rootNode)) posY = -50; /* **** quedé acá */

                node.Position = new Point(posX, posY);
                
                Canvas.SetLeft(node, node.Position.X);
                Canvas.SetTop(node, node.Position.Y);

                if (this.Children.Contains(node)) this.Children.Remove(node);
                this.Children.Add(node);
                return node.Position;
            }
        }

        private void DrawConnectionLines(Node node)
        {
            if (node.HasChildren && node.IsExpanded)
            {
                double startX = 0;
                double startY = 0;
                double endX = 0;
                double endY = 0;

                if (node.ChildNodes.Count == 1)
                {
                    startX = endX = node.Position.X + node.Width / 2;
                    startY = node.Position.Y + node.Height;
                    endY = startY + _vSpace;

                    Line line = GetConnectorLine(startX, endX, startY, endY);

                    this.Children.Add(line);
                    DrawConnectionLines(node.ChildNodes[0]);
                }
                else
                {
                    //first line --> from parent to middle line
                    startX = endX = node.Position.X + node.Width / 2;
                    startY = node.Position.Y + node.Height;
                    endY = startY + _vSpace / 2;

                    Line line = GetConnectorLine(startX, endX, startY, endY);
                    this.Children.Add(line);

                    //second line --> middle line
                    Node first = node.ChildNodes[0];
                    Node last = node.ChildNodes[node.ChildNodes.Count - 1];

                    startX = first.Position.X + first.Width / 2;
                    startY = endY = first.Position.Y - _vSpace / 2;
                    endX = last.Position.X + last.Width / 2;

                    line = GetConnectorLine(startX, endX, startY, endY);
                    this.Children.Add(line);

                    //draw line for every children
                    for (int i = 0; i <= (node.ChildNodes.Count - 1); i++)
                    {
                        Node current = node.ChildNodes[i];
                        startX = endX = current.Position.X + current.Width / 2;
                        endY = startY + _vSpace / 2;

                        line = GetConnectorLine(startX, endX, startY, endY);
                        this.Children.Add(line);
                        DrawConnectionLines(current);
                    }
                }
            }
        }

        /// <summary>
        /// Especifíca el estilo de la línea
        /// </summary>
        /// <param name="X1"></param>
        /// <param name="X2"></param>
        /// <param name="Y1"></param>
        /// <param name="Y2"></param>
        /// <returns>Line</returns>
        private Line GetConnectorLine(double X1, double X2, double Y1, double Y2)
        {
            Line line = new Line();
            line.StrokeThickness = 0.55;
            
            DoubleCollection dashArray = new DoubleCollection();
            dashArray.Add(1); //10 dash width
            dashArray.Add(2); //5 space betweem dashes
            
            line.StrokeDashArray = dashArray;
            
            line.Stroke = new SolidColorBrush(Colors.Black);

            /*nuevo pablo*/
            line.StrokeDashCap = PenLineCap.Round;
            line.StrokeStartLineCap = PenLineCap.Round;
            line.StrokeThickness = 3;

            line.Stroke = new SolidColorBrush(Colors.DarkGray);
            /*fin nuevo pablo*/

            line.X1 = X1;
            line.X2 = X2;
            line.Y1 = Y1;
            line.Y2 = Y2;

            return line;
        }
        #endregion
    }
}
