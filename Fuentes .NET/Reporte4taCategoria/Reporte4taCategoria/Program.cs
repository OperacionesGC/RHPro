using System;
using System.Collections.Generic;
using System.Diagnostics;
using RHPro.Repository;
using RHPro.Shared.Data;
using RHPro.Shared.IO;
using System.Data.SqlClient;
using RHPro.Shared.Ganancias.Data;
using RHPro.ADP.Data.Document;
using System.Text.RegularExpressions;

namespace Reporte4taCategoria
{
    class Program
    {
        static Log FLog;
        static batch_proceso eBProcess;
        static List<empleado> eBEmpleados;
        static UnitOfWork uof;
        static UnitOfWork uofProgress;
        static ContextReporte4taCategoria db;
        static ContextProgress dbProgress;
        static long TiempoAcumulado;
        static long TiempoInicialProceso;
        static string _connectionString;
        static string[] ArrParam;

        static int AcuSAC = 0;
        const decimal ERROR_REDONDEO = 0.01M;

        private static string removeProvider(string connectionString)
        {
            string PATTERN_CONTENT = @"((?i)provider.*?;)";
            Regex r = new Regex(PATTERN_CONTENT, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Singleline);
            Match m = r.Match(connectionString);
            if (m.Success)
            {
                return connectionString.Replace(m.Groups[1].Value, "").ToString();
            }
            else
            {
                return connectionString.ToString();
            }

        }

        static void Main(string[] args)
        {
            try
            {

#if DEBUG
                //args = new string[] { "418430", "LACAJA", "56238" };
                //args = new string[] { "791559", "Base_0_R3_ARG", "0" };
				//args = new string[] { "791579", "Base_0_R3_ARG", "0" };
				//args = new string[] { "47859", "RAET_AR_TEST", "0" };
				args = new string[] { "48932", "RAET_AR_TEST", "0" };
				


#endif

				Proceso eProcess = new Proceso(args, System.Reflection.Assembly.GetExecutingAssembly().Location);
                DataAcces Config = new DataAcces(AppDomain.CurrentDomain.BaseDirectory.Trim(), eProcess.Etiqueta, eProcess.Seed);


                FLog = new Log(Config.PathLog, "F6494taCategoria_" + eProcess.nroProceso + ".log", true, true, true);

                _connectionString = removeProvider(Config.Conexion);

                db = new ContextReporte4taCategoria(_connectionString);
                dbProgress = new ContextProgress(_connectionString);
                
                uofProgress = new UnitOfWork(dbProgress);
                uof = new UnitOfWork(db);

                FLog.WriteLine("", true);
                FLog.WriteLog(eProcess.Log);

                batch_procesoRepository rBatchProceso = new batch_procesoRepository(db);

                eBProcess = rBatchProceso.Find(eProcess.nroProceso);
                // uof.Repository<batch_proceso>().Find(eProcess.nroProceso);
                eBEmpleados = rBatchProceso.EmpleadosProceso(eProcess.nroProceso);
                

                eBProcess.bprcprogreso = 0;
                eBProcess.bprcestado = "Procesando";
                eBProcess.bprcfecInicioEj = DateTime.Now;
                eBProcess.bprctiempo = "0";
                eBProcess.bprcPid = Process.GetCurrentProcess().Id;

                uofProgress.Repository<batch_proceso>().Update(eBProcess);
                uofProgress.Save();

                FLog.WriteLine("\r");
                FLog.WriteLine("Cargando Parámetros del proceso.", false);

                ArrParam = eBProcess.bprcparam.Split('@');
                
                 
                if (!string.IsNullOrEmpty(ArrParam[0]))
                {
                    confrepRepository rConfrep = new confrepRepository(db);
                    List<confrep> eConfrep3 = rConfrep.GetConfrep(528);

                    foreach (confrep Confrep in eConfrep3)
                    {
                        if (Confrep.confnrocol == 1)
                        {
                            if (Confrep.conftipo == "ACM")
                            {
                                AcuSAC = Confrep.confval;
                            }
                        }
                    }

                        GenerarDatos();

                }
                else
                {
                    eBProcess.bprcFecFinEj = DateTime.Now;
                    eBProcess.bprcestado = "Abortado";
                    uofProgress.Repository<batch_proceso>().Update(eBProcess);
                    uofProgress.Save();

                    throw new Exception("Falla en los parametros del proceso.");
                }
            }
            catch (Exception e)
            {
                if (FLog != null)
                {
                    FLog.WriteLine("Mensaje: " + e.Message);
                    FLog.WriteLine("Excepción: " + ((e.InnerException != null) ? e.InnerException.Message.ToString() : ""));
                    FLog.WriteLine("Trace: " + ((e.StackTrace != null) ? e.StackTrace.ToString() : ""));

                    //Actualizo batch_Proceso
                    eBProcess.bprcFecFinEj = DateTime.Now;
                    eBProcess.bprcestado = "Error";
                    uofProgress.Repository<batch_proceso>().Update(eBProcess);
                    uofProgress.Save();
                }
            }
            finally
            {
                if (FLog != null)
                {
                    eBProcess.bprcFecFinEj = DateTime.Now;
                    eBProcess.bprcHoraFinEj = DateTime.Now.ToString("HH:mm:ss");
                    uofProgress.Repository<batch_proceso>().Update(eBProcess);
                    eBProcess.bprcestado = "Procesado";
                    eBProcess.bprcprogreso = 100;
                    batch_empleadoRepository rBatchEmpleado = new batch_empleadoRepository(db);
                    foreach (empleado emp in eBEmpleados)
                    {
                        //uofProgress.Repository<batch_empleado>().Delete(emp);
                        rBatchEmpleado.DeleteBatch_empleado(emp.ternro, eBProcess.bpronro);
                    }
                    
                    FLog.WriteLine("\rFin del procesamiento");
                    FLog.Flush();
                }
                uofProgress.Save();
            }
        }


    
        public static String BuscarEmpleados()
        {
            string lista = "0";
            foreach (empleado emp in eBEmpleados)
            {
                lista = lista + "," + emp.ternro;
            }
            return lista;
        }


        public static double BuscarSAC(double Ternro , int nrosac, int anio, int acunro )
        {
            SqlConnection cn = new SqlConnection(_connectionString);
            cn.Open();

            string StrSQL = " SELECT SUM(acu_mes.ammonto) ";
            StrSQL += " FROM acu_mes ";
            StrSQL += " WHERE  acu_mes.ternro = " + Ternro;
            StrSQL += " AND acunro=" + acunro;
            StrSQL += " AND amanio = " + anio;
            if (nrosac == 1)
            {
                StrSQL += " AND ammes <=6 ";
            }
            else
            {
                StrSQL += " AND ammes >= 7 ";
            }
           
            SqlCommand cmd = cn.CreateCommand();
            decimal SACacu = 0;
            cmd.CommandText = StrSQL;
            SqlDataReader reader = cmd.ExecuteReader();
            
            while (reader.Read())
            {
                if (!reader.IsDBNull(0))
                { SACacu = reader.GetFieldValue<decimal>(0); }
                else {
                    SACacu = 0;
                }
            }
            reader.Close();
            return (double)SACacu;

        }

        public static double BuscarSAC(double Ternro, int nrosac, string fecha, int acunro)
        {
			string FechaDesdeSac = "01/01/" +  DateTime.Parse(fecha).Year.ToString();
			
			int anio = DateTime.Parse(fecha).Year;

			SqlConnection cn = new SqlConnection(_connectionString);
            cn.Open();
			 
			SqlCommand cmd = cn.CreateCommand();

			string StrSQL = " SELECT acu_liq.almonto FROM acu_liq ";			
			StrSQL += " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro ";
			StrSQL += " INNER JOIN proceso ON proceso.pronro = cabliq.pronro ";
			StrSQL += " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro ";
			StrSQL += " WHERE Empleado = " + Ternro + " AND acunro = " + acunro ;

			if (nrosac == 1)
			{
				StrSQL += " AND (proceso.profecpago >= '01/01/" +  DateTime.Parse(fecha).Year.ToString() + "' AND proceso.profecpago <= '30/06/" + DateTime.Parse(fecha).Year.ToString() + "') ";
			}
			else
			{
				StrSQL += " AND (proceso.profecpago >= '01/07/" + DateTime.Parse(fecha).Year.ToString() + "' AND proceso.profecpago <= '31/12/" + DateTime.Parse(fecha).Year.ToString() + "') ";
			}			
			StrSQL += " AND (proceso.profecpago >= '" + FechaDesdeSac  + "' AND proceso.profecpago <= '" + fecha + "') ";


			cmd.CommandText = StrSQL;
			SqlDataReader reader = cmd.ExecuteReader();

			decimal SACacu = 0;
			while (reader.Read())
            {
                if (!reader.IsDBNull(0)) {				
						SACacu += reader.GetFieldValue<decimal>(0);
				}
   
            }
            reader.Close();
            return (double)SACacu;

        }

        public static void GenerarDatos()
        {
            try
            {

                SqlConnection cn = new SqlConnection(_connectionString);
                cn.Open();

                ganancias Ganancias;
                empleadoRepository empleadoRepository = new empleadoRepository(db);
                empleado eEmpleado;
                ter_docRepository DocumentoRepository = new ter_docRepository(db);

                ter_doc eCuilEmpleado;

                empresaRepository empreRepository = new empresaRepository(db);
                empresa eEmpresaEmpleado;
                ter_doc eCuitEmpresa;

                string sql = "";

                //emptip
                int emptip = int.Parse(ArrParam[0]);
                //estrnro empresa
                int empresa = int.Parse(ArrParam[1]);
                //fecha hasta
                string fechasta = ArrParam[2];
                //topedesde
                decimal topedesde = decimal.Parse(ArrParam[3]);
                //topehasta
                decimal topehasta = decimal.Parse(ArrParam[4]);
                //anual
                int anual = int.Parse(ArrParam[5]);
                //original
                int original = int.Parse(ArrParam[6]);
                //suscribe
                string suscribe = ArrParam[7];
                //caracter
                string caracter = ArrParam[8];
                //lugar
                string lugar = ArrParam[9];
                //fecha
                string FechaRev = ArrParam[10];
                //retiene
                int retenc = int.Parse(ArrParam[11]);

                //listempleados
                string listempleados = "";


                DateTime Fecha = DateTime.Parse(fechasta);
                int annio = Fecha.Year;

                string fechaInicio = "01/01/" + annio;
				
				int pronro;
                int ternro;
                double Progreso = 0;
                double IncPorc;
                //bool GetAceptaTope;
                string periodo = Fecha.Year.ToString();
                decimal RemBruta = 0;
                decimal RetnoHab = 0;
                decimal SAC1 = 0;
                decimal SAC2 = 0;
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
                decimal RemExHsExtras = 0;
                decimal OtrasDeduc = 0;
                decimal Alicuota = 0;
                decimal AlicuotaAplicable = 0;
                string fechaActual = Convert.ToDateTime(DateTime.Now).ToString("dd/MM/yyyy");

                //decimal rubro8 = 0;
                //decimal rubro9 = 0;
                decimal ImpDet = 0;
                decimal ImpRet = 0;
                decimal Pagacta = 0;
                decimal viaticos = 0;
                decimal alquileres = 0;
                //int empresaaux = 0;
                int contador = 0;
                string EmplNom = "";

                if (emptip != 3){
                    if (eBEmpleados.Count == 0)
                    {
                        FLog.WriteLine("No se encontraron empleados asociados al proceso.");
                        throw new Exception("No se encontraron empleados asociados al proceso.");

                    }
                    listempleados = BuscarEmpleados();
                }else
                {
                    FLog.WriteLine("Se procesarán todos los empleados...");
                }

                eEmpresaEmpleado = empreRepository.GetEmpresaByEstrnro(empresa);
                eCuitEmpresa = DocumentoRepository.GetDocument(eEmpresaEmpleado.ternro, 6);
               
                 SqlCommand cmd = cn.CreateCommand();
            

                Gan4taCab eGan4taCab;
                eGan4taCab = new Gan4taCab();
                eGan4taCab.bpronro = eBProcess.bpronro;
                eGan4taCab.fecha = DateTime.Parse(fechaActual);
                eGan4taCab.pronro = 0;
                eGan4taCab.EmprEstrnro = empresa;
                eGan4taCab.EmprNom = eEmpresaEmpleado.empnom;
                if (eCuitEmpresa != null)
                {
                    eGan4taCab.Emprcuit = eCuitEmpresa.nrodoc;
                }
                else {
                    eGan4taCab.Emprcuit = "0000000000";
                }
                eGan4taCab.Periodo = periodo.ToString();
                eGan4taCab.Original = original;
                eGan4taCab.TopeDesde = topedesde;
                eGan4taCab.TopeHasta = topehasta;
                eGan4taCab.Anual = anual;
                eGan4taCab.Caracter = caracter;
                eGan4taCab.Suscribe = suscribe;
                eGan4taCab.Lugar = lugar;
                eGan4taCab.FechaRev = DateTime.Parse(FechaRev);
                eGan4taCab.Retenc = retenc;
                
                eGan4taCab.detalle = new List<Gan4taDet>();
                uof.Repository<Gan4taCab>().Insert(eGan4taCab);
                //uof.Save();

                sql = " SELECT count(*) FROM  ( ";
                sql += " SELECT  traza_gan.ternro ternroInt, max(fecha_pago) fp FROM  traza_gan ";
                sql += " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro ";
                sql += " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 AND empresa.estrnro = " + empresa;
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
                if (emptip != 3)
                {
                    sql += " AND traza_gan.ternro in (" + listempleados + ")";
                    
                }
                sql += " group by  traza_gan.ternro) ganancia";

                cmd.CommandText = sql;
                int cant = int.Parse(cmd.ExecuteScalar().ToString());
                FLog.WriteLine("Cantidad de empleados alcanzados :" + cant);

                sql = " SELECT (select max(pronro) from traza_gan where fecha_pago = fp and ternro = ternroInt) pronro, ternroInt, fp, legajo FROM  ( ";
                sql += " SELECT  traza_gan.ternro ternroInt, max(fecha_pago) fp, empleado.empleg legajo FROM  traza_gan ";
                sql += " INNER JOIN empleado ON traza_gan.ternro = empleado.ternro ";
                sql += " INNER JOIN his_estructura  empresa ON empresa.ternro = traza_gan.ternro and empresa.tenro = 10 AND empresa.estrnro = " + empresa;
                sql += " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro ";
                sql += " INNER JOIN proceso ON proceso.pronro = traza_gan.pronro  ";
                sql += " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro ";
                sql += " WHERE traza_gan.fecha_pago <= '" + fechasta + "'" ;
                sql += " AND fecha_pago >= '" + fechaInicio + "'";
                sql += " AND (htetdesde <= '" + fechasta + "'";
                if (retenc == -1)
                {
                    sql += " AND retenciones>0";
                }
                sql += " ) AND ";
                sql += " (('" + fechasta + "'";
                sql += " <= htethasta) or (htethasta is null))";
    
                if (emptip != 3)
                {
                    sql += " AND traza_gan.ternro in (" + listempleados + ")";
                }
                sql += " group by  empleado.empleg, traza_gan.ternro) ganancia ORDER BY legajo, ternroInt ASC "; 

                cmd.CommandText = sql;
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    IncPorc = Math.Round(Convert.ToDouble(100) / Convert.ToDouble(cant), 4);
                    while (reader.Read())
                    {
                        pronro = reader.GetFieldValue<int>(0);
                        ternro = reader.GetFieldValue<int>(1);
					
						eCuilEmpleado = DocumentoRepository.GetDocument(ternro, 10);

                        Ganancias = new ganancias(db, pronro, ternro);
                           
                        if (Ganancias.aceptaTope(topedesde, topehasta))
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
                            HonServ= Math.Abs(Ganancias.montoItem(20));
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
                            //rubro8 = (decimal)Ganancias.rubro8();
                            //ImpDet = decimal.Round(rubro8,2);

							ImpDet = (decimal)Ganancias.rubro8();

							//rubro9 = Ganancias.rubro9();
                            Alicuota = (decimal)Ganancias.Alic();
                            AlicuotaAplicable= (decimal)Ganancias.AlicAplic();
                            SAC1 = (decimal)BuscarSAC(ternro, 1, fechasta, AcuSAC);
                            SAC2 = (decimal)BuscarSAC(ternro, 2, fechasta, AcuSAC);
							RemBruta = Ganancias.sumaMontoItems("1, 2, 34", true) - SAC1 - SAC2;

							ImpRet = Math.Abs(Ganancias.retenciones(DateTime.Parse(fechasta)));
                            Pagacta = Math.Abs(Ganancias.montoItem56());

                            viaticos = Ganancias.montoItem(30);
                            alquileres = Ganancias.montoItem(32);

                            Gan4taDet eGan4taDet = new Gan4taDet();

                            eGan4taDet.orden = contador;
                            eGan4taDet.Emplleg = eEmpleado.empleg;

                            EmplNom = eEmpleado.terape;

                            if (eEmpleado.terape2 != null){
                                EmplNom = EmplNom + " " + eEmpleado.terape2;
                            }

                            EmplNom = EmplNom + "," + eEmpleado.ternom;

                            if (eEmpleado.ternom2 != null){
                                EmplNom = EmplNom + " " + eEmpleado.ternom2;
                            }

                            if (eCuilEmpleado == null)
                            {
                                FLog.WriteLine("Empleado " + eEmpleado.empleg + " sin CUIL.");
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
						
						Progreso = Progreso + IncPorc;

                        TiempoAcumulado = Environment.TickCount;
                        eBProcess.bprcprogreso = Convert.ToDecimal(Progreso);
                        eBProcess.bprcPid = Process.GetCurrentProcess().Id;
                        uofProgress.Repository<batch_proceso>().Update(eBProcess);
                        uofProgress.Save();

                    }
                }
                reader.Close();
                uof.Save();

                
            }
            catch (Exception e)
            {
                Console.WriteLine("Could not connect to database: " + e.Message);
        
            }
            finally
            {

            }

        }


        
    }
}



