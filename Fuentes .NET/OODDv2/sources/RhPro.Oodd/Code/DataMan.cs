using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

public static class DataMan
{
    /// <summary>
    /// Corta el texto de ser necesario
    /// </summary>
    /// <param name="value">Texto a cortar</param>
    /// <returns>Texto cortado</returns>
    public static string cutWord(string value)
    {
        string val = value;

        if (value.Contains(" "))
            val = value.Substring(0, value.IndexOf(" ")).ToString();

        return val;
    }

    /// <summary>
    /// Permite que se escriban únicamente números en un TextBox
    /// </summary>
    /// <param name="e">EventArg de la tecla presionada</param>
    public static void allowNumbers(KeyEventArgs e)
    {
        if (e.PlatformKeyCode.Equals(Key.Back))
            return;

        double d = 0;
        string s = string.Empty;

        int c = e.PlatformKeyCode;

        if (c > 95 && c < 106) c -= 48;

        if (e.PlatformKeyCode != 190)
            s += Convert.ToChar(c);
        else
            s += ".0";

        if (c != 9)
            e.Handled = !double.TryParse(s, out d);
    }
}

