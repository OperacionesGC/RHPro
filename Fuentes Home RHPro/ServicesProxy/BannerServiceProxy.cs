using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.ar.com.rhpro.prueba;

namespace ServicesProxy
{
  public  static class BannerServiceProxy
    {
      private const string DataBaseDescFieldName = "hbandesc";
      private const string DataBaseImageFieldName = "hbanimage";
      private const string DataBaseDescextFieldName = "hbandescext";

      /// <summary>
      /// Busca banners 
      /// </summary>
      /// <param name="databaseID">Id de la base de datos</param>
      /// <param name="idioma">Idioma actual</param>
      /// <returns>Lista de banners</returns>
      public static List<Banner> Find(string databaseID, string idioma)
      {
          List<Banner> BannerToReturn = new List<Banner>();

          Consultas consultas = new Consultas();

          consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

          DataTable dataBases = consultas.Banner(databaseID, idioma);

          foreach (DataRow dataRow in dataBases.Rows)
          {
              Banner banner = new Banner
                                  {
                                      Description = dataRow[DataBaseDescFieldName].ToString(),
                                      ImageUrl = dataRow[DataBaseImageFieldName].ToString(),
                                      Titulo = dataRow[DataBaseDescextFieldName].ToString()

                                  };

              BannerToReturn.Add(banner);
          }

          return BannerToReturn;
      }   

    }
}
