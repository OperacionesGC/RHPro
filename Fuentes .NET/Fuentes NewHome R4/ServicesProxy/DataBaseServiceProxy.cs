using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Common;
using Entities;
using ServicesProxy.rhdesa;
using System.Net;

namespace ServicesProxy
{

    public static class DataBaseServiceProxy
    {
        private enum FieldPosition
        {
            Id = 1,
            Name = 0,
            IsDefaul = 3,
            IntegrateSecurity = 2
        }

        /// <summary>
        /// Busca las Base de Datos
        /// </summary>
        /// <returns>Lista de Base de Datos</returns>
        public static List<DataBase> Find(string dsm)
        {
            List<DataBase> dataBasesToReturn = new List<DataBase>();

            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.comboBase();
            

            if (dsm == "l")
            {
                DataView dv = dataBases.DefaultView;
                dv.Sort = "combo ASC";
                dataBases = dv.ToTable();
            }

            foreach (DataRow dataRow in dataBases.Rows)
            {
                string[] dataBaseArray = dataRow[0].ToString().Split(",".ToCharArray());

                // 4 Es la cantidad de parametros (FieldPosition) separados por "," esperados en cada row. 
                if (dataBaseArray.Count() >= 4)
                {
                    DataBase dataBase = new DataBase
                                            {
                                                Id = dataBaseArray[(int) FieldPosition.Id],
                                                Name = dataBaseArray[(int) FieldPosition.Name],
                                                IntegrateSecurity = (Utils.IntegrateSecurityConstants)(int.Parse(dataBaseArray[(int) FieldPosition.IntegrateSecurity])),
                                                IsDefault =(Utils.IsDefaultConstants)(int.Parse(dataBaseArray[(int) FieldPosition.IsDefaul]))
                                            };

                    dataBasesToReturn.Add(dataBase);
                 
                }
            }
           

            return dataBasesToReturn;
        }

    }
}
