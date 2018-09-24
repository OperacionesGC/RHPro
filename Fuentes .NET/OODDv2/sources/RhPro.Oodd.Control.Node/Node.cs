using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using RhPro.Oodd.Utilities;

namespace RhPro.Oodd.Control.Node
{
    #region Delegates
    public delegate void NodeSelectionChangedHandler(Node selectedNode);
    public delegate void NodeStateChangedHandler(Node currentNode, bool isExpanded);
    #endregion

    public class Node : Canvas, INotifyPropertyChanged
    {
        #region Public Events
        public event NodeSelectionChangedHandler NodeSelectionChanged;
        public event NodeStateChangedHandler NodeStateChanged;
        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        #region Private Members
        private bool _isExpanded;
        
        #endregion

        #region Private Const Members
        public double _width = 220;
        public double _height = 100; //75
        #endregion

        #region Public Properties
        public int ID { get; set; }
        public new string Legajo { get; set; }
        public new string Name { get; set; }
        public new string CompleteName { get; set; }
        public new string LastName { get; set; }
        public new string Division { get; set; } //departamento
        public new string DivisionDesc { get; set; } //departamentoDesc     //Lisandro Moro
        public new string Task { get; set; } //cargo    //Lisandro Moro
        public new string TaskDesc { get; set; } //cargoDesc
        public new string Phone { get; set; }
        public new string Email { get; set; }
        public new string HeadShot { get; set; } //url imagen
        public new string Company { get; set; }
        public new string CompanyDesc { get; set; }//EmpresaDesc    //Lisandro Moro
        public new string RealCompany { get; set; }
        public Node OldParentNode { get; set; }
        public Node ParentNode { get; set; }
        public bool HasChildren { get { return (ChildNodes != null && ChildNodes.Count > 0); } }
        public NodeCollection ChildNodes { get; set; }
        public NodeContent Content { get; set; }
        public Point Position { get; set; }
        public DragDropManager DragDrop { get; set; }
        public OrganizationChart OrgChart { get; set; }

        public bool IsExpanded
        {
            get { return _isExpanded; }
            set
            {
                _isExpanded = value;
                NotifyPropertyChanged("IsExpanded");
            }
        }
        #endregion

        #region Constructors
        public Node()
        {
            this.Width = _width;
            this.Height = _height;

            this.Background = new SolidColorBrush(Colors.Transparent);

            ChildNodes = new NodeCollection();
            Position = new Point();

            this.Loaded += OnLoaded;
            this.MouseLeftButtonDown += OnMouseLeftButtonDown;
            //this.MouseLeftButtonUp += OnMouseLeftButtonDown;
            this.PropertyChanged += OnPropertyChanged;
        }
        #endregion

        #region Event Handling
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (Content != null && !this.Children.Contains(Content))
            {
                //para re-dibujar el organigrama antes me fijo si tiene los datos antiguos y los borro
                if (this.Children.Count > 0) this.Children.Clear();
                this.Children.Add(Content);

                this.DragDrop.SetNodeContent(this.Content);
                //Content.BrdLowerRectangle.MaxWidth = 150;
                //DragDrop.RegisterDraggable(Content.StackPanelDrop);
                //DragDrop.RegisterDropTarget(Content.StackPanelDrag);
            }
        }

        private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Node selectedNode = sender as Node;

            if (selectedNode.OrgChart._selectedNode != null)
            {
                selectedNode.OrgChart._selectedNode.Content.Reset();

                if (!selectedNode.OrgChart._selectedNode.Equals(selectedNode))
                {
                    selectedNode.OrgChart._selectedNode = selectedNode;
                    selectedNode.OrgChart._selectedNode.Content.Select();
                }
                else
                    selectedNode.OrgChart._selectedNode = null;
            }
            else
            {
                selectedNode.OrgChart._selectedNode = selectedNode;
                selectedNode.OrgChart._selectedNode.Content.Select();
            }


        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (NodeStateChanged != null)
                NodeStateChanged(this, !_isExpanded);
        }

        private void NotifyPropertyChanged(String info)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(info));
        }
        #endregion

    }
}
