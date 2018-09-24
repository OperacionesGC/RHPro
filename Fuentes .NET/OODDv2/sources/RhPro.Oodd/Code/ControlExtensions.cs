using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Printing;

// ReSharper disable CheckNamespace - we are putting it into the same namespace as control to avoid using statements everywhere.
namespace System.Windows.Controls
// ReSharper restore CheckNamespace
{
    public static class ControlExtensions
    {
        private static readonly Point StandardLandscapePageSize = new Point(11, 8.5);
        private static readonly Point StandardLandscapePrintableAreaSize = new Point(10, 7.5);
        private const double DotsPerInch = 96;

        public static void PrintScreen(this StackPanel control)
        {
            try
            {
                double scaleX = 1;
                double scaleY = 1;
                double pixelWidth = DotsPerInch * StandardLandscapePrintableAreaSize.X;
                double pixelHeight = DotsPerInch * StandardLandscapePrintableAreaSize.Y;

                // determine scale ration to fit on a page
                if (control.ActualWidth > pixelWidth)
                    scaleX = pixelWidth / control.ActualWidth;

                if (control.ActualHeight > pixelHeight)
                    scaleY = pixelHeight / control.ActualHeight;

                //create scale transform to use the scale above to automatically resize the
                //control to be printed.
                var transform = new ScaleTransform
                {
                    CenterX = 0,
                    CenterY = 0,
                    ScaleX = scaleX,
                    ScaleY = scaleY
                };
                //create bit map for control, scaled appropriately
                var writableBitMap = new WriteableBitmap(control, transform);
                // put bit map on canvas
                var canvas = new Canvas
                {
                    Width = pixelWidth,
                    Height = pixelHeight,
                    Background = new ImageBrush { ImageSource = writableBitMap, Stretch = Stretch.Fill }
                };
                // create outer canvas to setup printable area margins
                var outerCanvas = new Canvas
                {
                    Width = StandardLandscapePageSize.X * DotsPerInch,
                    Height = StandardLandscapePageSize.Y * DotsPerInch
                };
                outerCanvas.Children.Add(canvas);
                //setup margins
                canvas.SetValue(Canvas.LeftProperty, DotsPerInch * (StandardLandscapePageSize.X - StandardLandscapePrintableAreaSize.X) / 2);
                canvas.SetValue(Canvas.TopProperty, DotsPerInch * (StandardLandscapePageSize.Y - StandardLandscapePrintableAreaSize.Y) / 2);
                //fore refresh just in case
                canvas.InvalidateMeasure();
                canvas.UpdateLayout();

                // create printable document
                var printDocument = new PrintDocument();
                printDocument.PrintPage += (s, args) =>
                {
                    args.PageVisual = outerCanvas;
                    args.HasMorePages = false;
                    
                };
                // launch print with the tile of Print Screen
                printDocument.Print("RHPro - Organigrama Dinámico");
            }
            catch (Exception exception)
            {
                // replace with real error handling
                MessageBox.Show("Un error ocurrió mientras se estaba imprimiendo. El mensaje de error es: " + exception.Message);
            }

        }

    }
}
