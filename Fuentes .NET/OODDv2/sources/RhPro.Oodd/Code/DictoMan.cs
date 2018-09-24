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
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace RhPro.Oodd.Code
{
    public static class DictoMan
    {
        public static IEnumerable SelectByKey(IDictionary<string,string> pDicto, string pKey)
        {
            IEnumerable ret = null;

            if(pDicto!=null)
                if(pDicto.Count > 0)
                    ret = pDicto.Select(x => x.Key.ToLower().Equals(pKey.ToLower()));

            return ret;
        }
    }
}
