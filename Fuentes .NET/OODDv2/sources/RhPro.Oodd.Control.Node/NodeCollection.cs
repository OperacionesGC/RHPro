using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace RhPro.Oodd.Control.Node
{
    public class NodeCollection : Collection<Node>
    {
        #region Public Events
        public event NodeSelectionChangedHandler NodeSelectionChanged;
        public event NodeStateChangedHandler NodeStateChanged;
        #endregion

        #region Constructors
        public NodeCollection(){}

        public NodeCollection(IEnumerable<Node> nodes)
        {
            AddRange(nodes);
        }
        #endregion

        #region Public Methods
        public void AddRange(IEnumerable<Node> values)
        {
            foreach (Node n in values)
                this.AddNode(n);
        }

        public void AddNode(Node item)
        {
            this.Add(item);
        }
        #endregion

        #region Public Functions
        /// <summary>
        /// Verifica que la colección no posea al nodo pasado por parámetro 
        /// dentro de todos sus nodos hijos hasta el nivel mas interior.
        /// </summary>
        /// <param name="item">Nodo</param>
        /// <returns>True si contiene el nodo.</returns>
        public bool ContainsNode(Node item)
        {
            if (this.Contains(item))
                return true;
            else
                foreach (Node n in this)
                    if (n.ChildNodes.ContainsNode(item))
                        return true;

            return false;
        }
        #endregion
    }
}
