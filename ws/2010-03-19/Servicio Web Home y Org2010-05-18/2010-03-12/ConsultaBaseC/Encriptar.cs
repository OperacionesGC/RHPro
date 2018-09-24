using System;
using System.Data;
using System.Configuration;


namespace ConsultaBaseC
{
    public class Encriptar
    {
        public static string Encrypt(string strEncryptionKey, string strTextToEncrypt)
        {
            string strTemp = "";
            for (int Outer = 0; Outer < strEncryptionKey.Length; Outer++)
            {
                int Key = (int)Convert.ToChar(strEncryptionKey.Substring(Outer, 1));
                for (int Inner = 0; Inner < strTextToEncrypt.Length; Inner++)
                {
                    strTemp = strTemp + Convert.ToString((char)(((int)Convert.ToChar(strTextToEncrypt.Substring(Inner, 1))) ^ Key));
                    Key = (Key + strEncryptionKey.Length) % 256;
                }
                strTextToEncrypt = strTemp;
                strTemp = "";
            }

            return CadenaHex(strTextToEncrypt);
        }

        public static string Decrypt(string strEncryptionKey, string strTextToEncrypt)
        {
            string strTemp = "";
            strTextToEncrypt = CadenaAscii(strTextToEncrypt);
            for (int Outer = 0; Outer < strEncryptionKey.Length; Outer++)
            {
                int Key = (int)Convert.ToChar(strEncryptionKey.Substring(Outer, 1));
                for (int Inner = 0; Inner < strTextToEncrypt.Length; Inner++)
                {
                    strTemp = strTemp + Convert.ToString((char)(((int)Convert.ToChar(strTextToEncrypt.Substring(Inner, 1))) ^ Key));
                    Key = (Key + strEncryptionKey.Length) % 256;
                }
                strTextToEncrypt = strTemp;
                strTemp = "";
            }

            return strTextToEncrypt;
        }

        public static string CadenaHex(string strTextToEncrypt)
        {
            string Buffer = "";
            for (int Outer = 0; Outer < strTextToEncrypt.Length; Outer++)
            { 
                strTextToEncrypt.Substring(Outer, 1);
                string Auxi = ((int)Convert.ToChar(strTextToEncrypt.Substring(Outer, 1))).ToString("X");
                if (Auxi.Length < 2)
                    Auxi = "0" + Auxi;
                Buffer = Buffer + Auxi;
            }
            return Buffer;
        }

        public static string CadenaAscii(string strTextToEncrypt)
        {
            string Buffer = "";
            for (int Outer = 0; Outer < strTextToEncrypt.Length; Outer = Outer + 2 )
            {
                Buffer = Buffer + Convert.ToString((char)int.Parse(strTextToEncrypt.Substring(Outer, 2), System.Globalization.NumberStyles.HexNumber));
            }
            return Buffer;
        }
    }


}
