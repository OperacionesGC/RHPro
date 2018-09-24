using System;
using System.Collections.Generic;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace RhPro.Oodd.Utilities
{
    /// <summary>
    /// Establece un contrato específico para la página principal del organigrama
    /// </summary>
    public interface IODPage
    {
        /// <summary>
        /// Habilita el boton guardar
        /// </summary>
        bool EnableSaveButton { get; set; }

        /// <summary>
        /// Habilita el botón Undo
        /// </summary>
        bool EnableUndoButton { get; set; }

        /// <summary>
        /// Agrega una nueva referencia de nodo a la colección Undo de tipo LIFO 
        /// </summary>
        /// <param name="undoNode">La referencia de nodo a agregar</param>
        void AddUndoNode(RevertReferenceNode undoNode);

        /// <summary>
        /// Agrega una nueva referencia de nodo a la colección Redo de tipo LIFO
        /// </summary>
        /// <param name="undoNode">La referencia de nodo a agregar</param>
        void AddRedoNode(RevertReferenceNode redoNode);

        /// <summary>
        /// Ruta de imagenes de los empleados
        /// </summary>
        string ImagePath{ get;}

        /// <summary>
        /// Especifica si se muestra el numero de legajo en los nodos
        /// </summary>
        string ShowDock { get; }

        /// <summary>
        /// Nivel de profundidad de undo/redo
        /// </summary>
        int MaxUndoLevel { get; }

        /// <summary>
        /// Indica si se muestra el nombre completo también en el Nodo
        /// </summary>
        bool ShowCompleteName { get; }
    }

    /// <summary>
    /// Guarda las referencias de un nodo, a fin de implementar los comandos Undo/Redo
    /// </summary>
    public class RevertReferenceNode
    {
        /// <summary>
        /// Identificador del Nodo
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Identificador del Nodo Padre anterior
        /// </summary>
        public int IdOldParent { get; set; }

        /// <summary>
        /// Identificador del Nodo Padre actual
        /// </summary>
        public int IdNewParent { get; set; }
    }
}
