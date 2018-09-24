using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;
using System.Diagnostics;
using System.Windows.Threading;
using RhPro.Oodd.Utilities;


//////////////////////////////////////////////////////////////////
//                                                              //
//This code is provided As-Is and confers no rights.            //
//                                                              //
//////////////////////////////////////////////////////////////////

namespace RhPro.Oodd.Control.Node
{
    //This class allows UIElements to be dragged around and to be dropped into another Panel.
    //Elements which need to able to be dragged should be registered as draggable by calling the 
    //RegisterDraggable method. Panels which are to have elements dragged into them should be
    //registered as drop targets by calling RegisterDropTarget.
    //As it stands, there are a number of limitations with this implementation:
    //1. The logic does not take into account RenderTransforms applied to either elements or Panels.
    //2. If multiple drop target panels overlap at a point where an element is dropped, the DragDropManager
    //2. Si múltiples paneles drop target convergen en un punto donde un elemento es colocado, el sistema
    //   will drop the element into the first Panel it finds, which may or may not be the Panel which is 
    //   rendered above the others.
    //3. Most Controls will be unusable while they are registered as draggable, since the 
    //   MouseLeftButtonDown event handler removes the element from the tree and re-adds it. Controls 
    //   should be unregistered for normal use.
    //4. All draggable elements must be inside Panels (not UserControls, Broders, ListBoxes, etc).
    //All of the above limitations could be fixed reasonably easily with modifications to be below code.
    public class DragDropManager
    {
        //A list of Panels which can have elements dragged into them
        public List<StackPanelNode> dropTargets = new List<StackPanelNode>();

        //This Canvas is used to hold elements while they are being dragged.
        private Canvas topCanvas;

        //This is the point on the element by which it is being dragged
        private Point grabPoint;

        //if the element is not dragged into a valid drop target, then it is
        //placed back to where it was originally. These variables are required
        //to restore the element to its previous location.
        private double backupCanvasLeft, backupCanvasTop;
        private int backupChildIndex;
        private StackPanelNode backupParent;

        private NodeContent _nodeContent;

        public void SetNodeContent(NodeContent nodeContent)
        {
            this._nodeContent = nodeContent;

            

            this.RegisterDraggable(_nodeContent.StackPanelDrag);
            this.RegisterDropTarget(_nodeContent.StackPanelDrop);

        }

        public DragDropManager(Canvas topCanvas)
        {
            this.topCanvas = topCanvas;
        }

        public void RegisterDraggable(StackPanelNode element)
        {
            element.MouseLeftButtonDown += ElementMouseLeftButtonDown;
        }

        public void UnRegisterDraggable(StackPanelNode element)
        {
            element.MouseLeftButtonDown -= ElementMouseLeftButtonDown;
        }

        public void RegisterDropTarget(StackPanelNode panel)
        {
            dropTargets.Add(panel);
        }

        public void UnRegisterDropTarget(StackPanelNode panel)
        {
            if (dropTargets.Contains(panel))
            {
                dropTargets.Remove(panel);
            }
        }

        //Start draqgging the element
        private void ElementMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            StackPanelNode element = sender as StackPanelNode;
            StackPanelNode parent = element.Parent as StackPanelNode;

            if (parent != null && element.CaptureMouse())
            {
                element.MouseLeftButtonUp += ElementMouseLeftButtonUp;
                element.MouseMove += ElementMouseMove;

                backupChildIndex = parent.Children.IndexOf(element);
                backupCanvasLeft = Canvas.GetLeft(element);
                backupCanvasTop = Canvas.GetTop(element);
                backupParent = parent;

                parent.Children.Remove(element);
                topCanvas.Children.Add(element);
                topCanvas.UpdateLayout();

                grabPoint.X = element.ActualWidth / 2;
                grabPoint.Y = element.ActualHeight / 2;

                UpdateElementPosition(element, e.GetPosition(topCanvas));
            }
            /*
            Node selectedNode = element.NodeOwner;

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
            */
        }

        private void ElementMouseMove(object sender, MouseEventArgs e)
        {
            Point point = e.GetPosition(topCanvas);
            UpdateElementPosition(sender as StackPanelNode, point);
        }

        private void UpdateElementPosition(StackPanelNode element, Point mousePoint)
        {
            /* Agregado posteriormente puede perjudicar la performance del arrastre */
            NodeContent nodeContent = element.NodeOwner.Content;
            SolidColorBrush solidColorBrush = new SolidColorBrush(nodeContent._selectedElement);

            if (!nodeContent.brdMain.Background.Equals(solidColorBrush))
                nodeContent.brdMain.Background = solidColorBrush;

            //Node selectedNode = element.NodeOwner;
            /*

            

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

            */



            /* Fin del agregado */

            Canvas.SetLeft(element, mousePoint.X - grabPoint.X);
            Canvas.SetTop(element, mousePoint.Y - grabPoint.Y);
        }

        //End the dragging operation.
        private void ElementMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            StackPanelNode element = sender as StackPanelNode;
            element.ReleaseMouseCapture();
            element.MouseLeftButtonUp -= ElementMouseLeftButtonUp;
            element.MouseMove -= ElementMouseMove;

            //Try to find a Panel to insert the element into.

            //The following code is the source of limitation #2 listed above.
            //Instead of checking the ActualWidth and ActualHeight of the Panel,
            //the UIElement.HitTest method could be used, which can be used to determine
            //which Panel is above all others at the specified point.
            //However, this method would require all targets to have a non-null Background set;
            //the hit-test works against the visuals in the tree.
            bool foundTargetPanel = false;
            int index = 0;
            while (index < dropTargets.Count && !foundTargetPanel)
            {
                StackPanelNode panel = dropTargets[index] as StackPanelNode;
                //StackPanelNode panel = (StackPanelNode)dropTargets[index];
                //Get the mouse position relative to the Panel:
                
                //panel.UpdateLayout();

                /*System.Windows.Threading.Dispatcher.BeginInvoke(delegate
                {*/
                    Point mousePos = e.GetPosition(panel);//.NodeOwner);

                //});
                //Point mousePos = e.GetPosition(null);//.NodeOwner);
                
                if (mousePos.X >= 0 && mousePos.Y >= 0
                    && mousePos.X <= panel.ActualWidth && mousePos.Y <= panel.ActualHeight)
                {
                    //se evita que el mismo objeto arrastrado se coloque a sí mismo en su mismo lugar
                    if (!panel.NodeOwner.Equals(element.NodeOwner))
                    {
                        topCanvas.Children.Remove(element);
                        AddElementToPanel(element, panel, mousePos);
                        foundTargetPanel = true;
                    }
                    //se agregó para mayor eficiencia, CHEQUEAR si está bien
                    break;
                }

                index++;
            }

            //if valid drop target Panel is not found, restore the element to its original location
            if (!foundTargetPanel)
            {
                topCanvas.Children.Remove(element);
                Canvas.SetLeft(element, backupCanvasLeft);
                Canvas.SetTop(element, backupCanvasTop);
                backupParent.Children.Insert(backupChildIndex, element);
            }
        }

        private void AddElementToPanel(StackPanelNode element, StackPanelNode panel, Point point)
        {
            if (panel is StackPanelNode)
            {
                AddElementToStackPanel(element, panel as StackPanelNode, point);
            }
        }

        private void AddElementToStackPanel(StackPanelNode element, StackPanelNode stackPanel, Point point)
        {
            //Determine the correct index to insert the element into the StackPanel
            int index = 0;
            if (stackPanel.Orientation == Orientation.Vertical)
            {
                double y = 0;
                while (index < stackPanel.Children.Count)
                {
                    var child = stackPanel.Children[index];
                    var slot = LayoutInformation.GetLayoutSlot(child as FrameworkElement);
                    y = slot.Bottom - (child as FrameworkElement).Margin.Bottom;
                    if (y > point.Y)
                    {
                        break;
                    }
                    index++;
                }
            }
            else
            {
                double x = 0;
                while (index < stackPanel.Children.Count)
                {
                    var child = stackPanel.Children[index];
                    var slot = LayoutInformation.GetLayoutSlot(child as FrameworkElement);
                    x = slot.Right - (child as FrameworkElement).Margin.Right;
                    if (x > point.X)
                    {
                        break;
                    }
                    index++;
                }
            }

            //se evita que el mismo objeto arrastrado se coloque a sí
            //mismo en su mismo lugar
            if (!element.NodeOwner.Equals(stackPanel.NodeOwner))
            {
                Node dragNode = element.NodeOwner;
                Node newParentNode = stackPanel.NodeOwner;                
                
                //verifica que el dragNode no sea padre del nuevo padre
                if (!dragNode.ChildNodes.ContainsNode(newParentNode))
                {
                    Node oldParentNode = dragNode.ParentNode;
                    oldParentNode.ChildNodes.Remove(dragNode);

                    dragNode.OldParentNode = oldParentNode;
                    dragNode.ParentNode = newParentNode;
                    newParentNode.ChildNodes.Add(dragNode);

                    newParentNode.OrgChart.RootVisual.EnableUndoButton = true;
                    newParentNode.OrgChart.RootVisual.AddUndoNode(new RevertReferenceNode{
                        Id=dragNode.ID, IdNewParent=newParentNode.ID, 
                        IdOldParent=oldParentNode.ID});
                }
                newParentNode.DragDrop.dropTargets.Clear();

                newParentNode.OrgChart.FillChart();

                newParentNode.OrgChart.RootVisual.EnableSaveButton = true;

                //newParentNode.OrgChart.UpdateLayout();
                
            }
        }

    }
}
