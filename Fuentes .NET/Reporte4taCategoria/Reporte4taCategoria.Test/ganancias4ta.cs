using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RHPro.Shared.Ganancias.Data;
using RHPro.Shared.Data;
using Reporte4taCategoria;
using System.Data.SqlClient;
using RHPro.ADP.Data.Document;
using System.Collections.Generic;
using RHPro.Repository;

namespace Reporte4taCategoria
{
    [TestClass]
    public class Reporte4taCategoria
    {
        ContextReporte4taCategoria db;
        string conexion = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=BASE_0_R3_ARG;Data Source=RHDESA";
        SqlCommand cmd;
        SqlConnection cn;
        UnitOfWork uof;
        const decimal ERROR_REDONDEO = 0.01M;

        private void connect() {
            db = new ContextReporte4taCategoria(conexion);
            cn = new SqlConnection(conexion);
            cmd = cn.CreateCommand();
            cn.Open();
            uof = new UnitOfWork(db);
        }



        private void gananciaAnio(string anio, int retenc)
        {
            string fechasta = "31/12/"+anio;
            string fechaInicio = "01/01/"+anio;

            int pronro;
            int ternro;
            string periodo = "2016";
            decimal RemBruta = 0;
			decimal RetnoHab = 0;
			decimal RemNoAlc = 0;
            decimal RemExc = 0;
            decimal RemOEmp = 0;
            decimal Apfondos = 0;
            decimal ApOS = 0;
            decimal CuotSin = 0;
            decimal ApjubOE = 0;
            decimal ApOSOE = 0;
            decimal CuotSinOE = 0;
            decimal CoutMedAsis = 0;
            decimal PrimSeg = 0;
            decimal GastSep = 0;
            decimal GastEst = 0;
            decimal DonFisc = 0;
            decimal DescObl = 0;
            decimal HonServ = 0;
            decimal IntCredH = 0;
            decimal ApCapSoc = 0;
            decimal EmpSerDom = 0;
            decimal ApCajCom = 0;
            decimal GanNI = 0;
            decimal DedEsp = 0;
            decimal CargFam = 0;
            decimal conyuge = 0;
            decimal Hijos = 0;
            decimal OtrasCar = 0;
			decimal Alicuota = 0;
			decimal AlicuotaAplicable = 0;
			decimal SAC1 = 0;
			decimal SAC2 = 0;

			string fechaActual = Convert.ToDateTime(DateTime.Now).ToString("dd/MM/yyyy");

            decimal rubro8 = 0;
            decimal rubro9 = 0;
            decimal ImpDet = 0;
            decimal ImpRet = 0;
            decimal Pagacta = 0;
			decimal RemExHsExtras = 0;
			decimal OtrasDeduc = 0;
			decimal viaticos = 0;
			decimal alquileres = 0;
			//int empresaaux = 0;
			int contador = 0;
            string EmplNom = "";


            ganancias Ganancias;
            empleadoRepository empleadoRepository = new empleadoRepository(db);
            empleado eEmpleado;
            ter_docRepository DocumentoRepository = new ter_docRepository(db);

            ter_doc eCuilEmpleado;

            empresaRepository empreRepository = new empresaRepository(db);

            //string sql = "select top 10000 cabliq.pronro,cabliq.empleado,proceso.profecini from cabliq inner join proceso on cabliq.pronro=proceso.pronro order by cabliq.pronro desc";

            string sql = " SELECT (select max(pronro) from traza_gan where fecha_pago = fp and ternro = ternroInt) pronro, ternroInt, fp, legajo FROM  ( ";
            sql += " SELECT  traza_gan.ternro ternroInt, max(fecha_pago) fp, empleado.empleg legajo FROM  traza_gan ";
            sql += " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro ";
            sql += " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 ";
            sql += " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro ";
            sql += " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro  ";
            sql += " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro ";
            sql += " WHERE traza_gan.fecha_pago <= '" + fechasta + "'";
            sql += " AND fecha_pago >= '" + fechaInicio + "'";
            sql += " AND (htetdesde <= '" + fechasta + "'";
            if (retenc == -1)
            {
                sql += " AND retenciones>0";
            }
            sql += " ) AND ";
            sql += " (('" + fechasta + "'";
            sql += " <= htethasta) or (htethasta is null))";

            sql += " AND traza_gan.ternro in (select ternro from empleado)";
      
            sql += " group by  empleado.empleg, traza_gan.ternro) ganancia ORDER BY legajo, ternroInt ASC ";


            cmd.CommandText = sql;
            SqlDataReader reader = cmd.ExecuteReader();
            Gan4taCab eGan4taCab;
            eGan4taCab = new Gan4taCab();
            eGan4taCab.bpronro = 99999;
            eGan4taCab.fecha = DateTime.Parse(fechaActual);
            eGan4taCab.pronro = 0;
            eGan4taCab.EmprEstrnro = 10;
            eGan4taCab.EmprNom = "Unit test";
            eGan4taCab.Emprcuit = "Unit test";
            eGan4taCab.Periodo = periodo.ToString();
            eGan4taCab.Original = -1;
            eGan4taCab.TopeDesde = 2;
            eGan4taCab.TopeHasta = 2;
            eGan4taCab.Anual = -1;
            eGan4taCab.Caracter = "Unit test";
            eGan4taCab.Suscribe = "Unit test";
            eGan4taCab.Lugar = "Unit test";
            eGan4taCab.FechaRev = DateTime.Now;
            eGan4taCab.Retenc = 0;

            eGan4taCab.detalle = new List<Gan4taDet>();
            uof.Repository<Gan4taCab>().Insert(eGan4taCab);

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    pronro = reader.GetFieldValue<int>(0);
                    ternro = reader.GetFieldValue<int>(1);
                    
                    eCuilEmpleado = DocumentoRepository.GetDocument(ternro, 10);

                    Ganancias = new ganancias(db, pronro, ternro);

                    if (Ganancias.aceptaTope(0, 1000000))
                    {
                        contador++;

                        eEmpleado = empleadoRepository.GetEmpleado(ternro);
                        eCuilEmpleado = DocumentoRepository.GetDocument(ternro, 10);

						RetnoHab = Math.Abs(Ganancias.montoItem(3));
						RemNoAlc = Ganancias.sumaMontoItems("19, 28", true);
						RemExc = 0;
						RemOEmp = Math.Abs(Ganancias.montoItem(4));
						Apfondos = Ganancias.sumaMontoItems("5, 35", true);
						ApOS = Ganancias.sumaMontoItems("6, 36", true);
						CuotSin = Ganancias.sumaMontoItems("7, 37", true);
						ApjubOE = Math.Abs(Ganancias.montoItem(25));
						ApOSOE = Math.Abs(Ganancias.montoItem(26));
						CuotSinOE = Math.Abs(Ganancias.montoItem(27));
						CoutMedAsis = Math.Abs(Ganancias.montoItem(13));
						PrimSeg = Math.Abs(Ganancias.montoItem(8));
						GastSep = Math.Abs(Ganancias.montoItem(9));
						GastEst = Ganancias.sumaMontoItems("23", true);
						DonFisc = Math.Abs(Ganancias.montoItem(15));
						DescObl = 0;   // ??????
						HonServ = Math.Abs(Ganancias.montoItem(20));
						IntCredH = Math.Abs(Ganancias.montoItem(21));
						ApCapSoc = Ganancias.sumaMontoItems("22,24", true);
						EmpSerDom = Math.Abs(Ganancias.montoItem(31));
						//ApCajCom = Math.Abs(Ganancias.montoItem(18));
						GanNI = Math.Abs(Ganancias.montoItem(17));
						DedEsp = Math.Abs(Ganancias.montoItem(16));
						CargFam = Ganancias.sumaMontoItems("10,11,12", true);
						conyuge = Math.Abs(Ganancias.montoItem(10));
						Hijos = Math.Abs(Ganancias.montoItem(11));
						OtrasCar = Math.Abs(Ganancias.montoItem(12));
						RemExHsExtras = Ganancias.montoItem(33);
						OtrasDeduc = Math.Abs(Ganancias.montoItem(18));
						rubro8 = (decimal)Ganancias.rubro8();
						ImpDet = decimal.Round(rubro8, 2);
						rubro9 = Ganancias.rubro9();
						Alicuota = (decimal)Ganancias.Alic();
						AlicuotaAplicable = (decimal)Ganancias.AlicAplic();

						SAC1 = (decimal)100;
						SAC2 = (decimal)200;
						RemBruta = Ganancias.sumaMontoItems("1, 2, 34", true) - SAC1 - SAC2;
						
						ImpRet = Math.Abs(Ganancias.retenciones(DateTime.Parse(fechasta)));
						Pagacta = Math.Abs(Ganancias.montoItem56());

						viaticos = Ganancias.montoItem(30);
						alquileres = Ganancias.montoItem(32);

						Gan4taDet eGan4taDet = new Gan4taDet();

						eGan4taDet.orden = contador;
						eGan4taDet.Emplleg = eEmpleado.empleg;

						EmplNom = eEmpleado.terape;

						if (eEmpleado.terape2 != null)
						{
							EmplNom = EmplNom + " " + eEmpleado.terape2;
						}

						EmplNom = EmplNom + "," + eEmpleado.ternom;

						if (eEmpleado.ternom2 != null)
						{
							EmplNom = EmplNom + " " + eEmpleado.ternom2;
						}

						if (eCuilEmpleado == null)
						{
							eGan4taDet.EmplCuit = "0";
						}
						else
						{
							if (eCuilEmpleado.nrodoc != "")
							{ eGan4taDet.EmplCuit = eCuilEmpleado.nrodoc; }
							else
							{ eGan4taDet.EmplCuit = "0"; }

						}
						eGan4taDet.EmplNom = EmplNom;
						eGan4taDet.Rembruta = RemBruta;
						eGan4taDet.RemNoAlc = RemNoAlc;
						eGan4taDet.RemExc = RemExc;
						eGan4taDet.RemOEmp = RemOEmp;
						eGan4taDet.Apfondos = Apfondos;
						eGan4taDet.ApOS = ApOS;
						eGan4taDet.CuotSin = CuotSin;
						eGan4taDet.ApjubOE = ApjubOE;
						eGan4taDet.ApOSOE = ApOSOE;
						eGan4taDet.CuotSinOE = CuotSinOE;
						eGan4taDet.CoutMedAsis = CoutMedAsis;
						eGan4taDet.PrimSeg = PrimSeg;
						eGan4taDet.GastSep = GastSep;
						eGan4taDet.GastEst = GastEst;
						eGan4taDet.DonFisc = DonFisc;
						eGan4taDet.DescObl = DescObl;
						eGan4taDet.HonServ = HonServ;
						eGan4taDet.IntCredH = IntCredH;
						eGan4taDet.ApCapSoc = ApCapSoc;
						eGan4taDet.EmpSerDom = EmpSerDom;
						eGan4taDet.ApCajCom = ApCajCom;
						eGan4taDet.GanNI = GanNI;
						eGan4taDet.DedEsp = DedEsp;
						eGan4taDet.CargFam = CargFam;
						eGan4taDet.conyuge = conyuge;
						eGan4taDet.Hijos = Hijos;
						eGan4taDet.OtrasCar = OtrasCar;
						eGan4taDet.ImpDet = ImpDet;
						eGan4taDet.ImpRet = ImpRet;
						eGan4taDet.Pagacta = Pagacta;
						eGan4taDet.alquileres = alquileres;
						eGan4taDet.viaticos = viaticos;
						eGan4taDet.RetnoHab = RetnoHab;
						eGan4taDet.RemExHsExtras = RemExHsExtras;
						eGan4taDet.SAC1 = SAC1;
						eGan4taDet.SAC2 = SAC2;
						eGan4taDet.OtrasDeduc = OtrasDeduc;
						eGan4taDet.Alicuota = Alicuota;
						eGan4taDet.AlicuotaAplicable = AlicuotaAplicable;
						eGan4taDet.fecPago = reader.GetFieldValue<DateTime>(2);

						eGan4taCab.detalle.Add(eGan4taDet);
						uof.Repository<Gan4taDet>().Insert(eGan4taDet);
					}
                }
            }
            reader.Close();
        }

        private void deleteOldData() {
            string sql = "delete from Gan4taDet where gan4tanro in (select gan4tanro from Gan4taCab where bpronro=99999); ";
            sql += "delete from Gan4taCab where bpronro = 99999;";
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
        }

        [TestMethod]
        public void ItemsGananciasSRetenciones()
        {
            string[] anios = { "2014" ,"2015", "2016", "2017" };

            connect();

            foreach (var anio in anios)
            {
                deleteOldData();
                gananciaAnio(anio, 0);
                uof.Save();
            }

            Assert.AreEqual(true, true);
        }

        [TestMethod]
        public void ItemsGananciasCRetenciones()
        {
			string[] anios = { "2014", "2015", "2016", "2017" };
			connect();

            foreach (var anio in anios)
            {
                deleteOldData();
                gananciaAnio(anio, -1);
                uof.Save();
            }

            Assert.AreEqual(true, true);
        }
    }
}
