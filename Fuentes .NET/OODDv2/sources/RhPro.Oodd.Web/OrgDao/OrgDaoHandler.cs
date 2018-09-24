using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Data;
using System.Xml.Linq;
using System.ServiceModel.Description;
using System.Configuration;
using RhPro.Oodd.Web.OrgService;
using RhPro.Oodd.Web.Services;
using RhPro.Oodd.Web.DataModel;
using System.Collections.Specialized;
using System.Diagnostics;

namespace RhPro.Oodd.Web.OrgDao
{
    public class OrgDaoHandler
    {
        public static int BASE;

        //Build the orgchart from a specific Employee
        public static OrgChart ReadOrgChart(long legajo, int maxLevel)
        {
            OrgChart orgChart = new OrgChart();
            orgChart.orgChartCode = 0;

            ConsultasSoapClient proxy = ServiceMan.Get();
            
            try
            {

                // Lisandro Moro - Validacion certificados https - SSL//
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) =>
                {
                    return true;
                };

                // Get the root employee to build the chart
                DataTable rootEmp = proxy.BuscarEmpleado(legajo, 1, BASE);
                
                if (rootEmp.Rows.Count > 0)
                {
                    DataRow row = rootEmp.Rows[0];
                    long codEmp = long.Parse(row["CodEmp"].ToString());
                    XElement tree = new XElement("root");
                    tree = GetEmployeeData(codEmp, proxy, tree);

                    // Build the Tree
                    int level = 1;
                    tree = BuildOrgChart(tree, proxy, level, maxLevel);

                    orgChart.returnCode = 0;
                    orgChart.tree = tree;
                }
                else
                {
                    Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChart : " + "El empleado: " + legajo + 
                        " no existe.", System.Diagnostics.EventLogEntryType.Error);

                    orgChart.returnCode = -2;
                    orgChart.errorMessage = "El empleado: " + legajo + " no existe.";
                }
            }
            catch (Exception ex)
            {
                Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChart : " + "Error al leer el organigrama : " + ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
                orgChart.returnCode = -1;
                orgChart.errorMessage = "No se pudo generar el organigrama, consulte al Administrador.";
            }

            //Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChart : " + "Sali", System.Diagnostics.EventLogEntryType.Information);

            return orgChart;
        }

        //Build the orgchart from the next Employee
        public static OrgChart ReadOrgChartFromNextEmp(long legajo, int maxLevel)
        {
            OrgChart orgChart = new OrgChart();
            orgChart.orgChartCode = 0;

            Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromNextEmp : " + 
                "Entre", System.Diagnostics.EventLogEntryType.Information);

            ConsultasSoapClient proxy = ServiceMan.Get();

            try
            {
                // Get the root employee to build the chart
                DataTable rootEmp = proxy.SgtEmpl(legajo, 1, BASE);
                if (rootEmp.Rows.Count > 0)
                {
                    DataRow row = rootEmp.Rows[0];
                    long codEmp = long.Parse(row["CodEmp"].ToString());
                    XElement tree = new XElement("root");
                    tree = GetEmployeeData(codEmp, proxy, tree);

                    // Build the Tree
                    int level = 1;
                    tree = BuildOrgChart(tree, proxy, level, maxLevel);

                    orgChart.returnCode = 0;
                    orgChart.tree = tree;
                }
                else
                {
                    Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromNextEmp : " + "Es el último empleado.", System.Diagnostics.EventLogEntryType.Error);
                    orgChart.returnCode = -4;
                    orgChart.errorMessage = "Es el último empleado.";
                }
            }
            catch (Exception ex)
            {
                Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromNextEmp : " + "Error al leer el organigrama : " + ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
                orgChart.returnCode = -1;
                orgChart.errorMessage = "No se pudo generar el organigrama, consulte al Administrador.";
            }

            Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromNextEmp : " + "Sali", System.Diagnostics.EventLogEntryType.Information);

            return orgChart;
        }

        //Build the orgchart from the previous Employee
        public static OrgChart ReadOrgChartFromPreviousEmp(long legajo, int maxLevel)
        {
            OrgChart orgChart = new OrgChart();
            orgChart.orgChartCode = 0;

            Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromPreviousEmp : " 
                + "Entre", System.Diagnostics.EventLogEntryType.Information);

            ConsultasSoapClient proxy = ServiceMan.Get();

            try
            {
                // Get the root employee to build the chart
                DataTable rootEmp = proxy.AntEmpl(legajo, 1, BASE);
                if (rootEmp.Rows.Count > 0)
                {
                    DataRow row = rootEmp.Rows[0];
                    long codEmp = long.Parse(row["CodEmp"].ToString());
                    XElement tree = new XElement("root");
                    tree = GetEmployeeData(codEmp, proxy, tree);

                    // Build the Tree
                    int level = 1;
                    tree = BuildOrgChart(tree, proxy, level, maxLevel);

                    orgChart.returnCode = 0;
                    orgChart.tree = tree;
                }
                else
                {
                    Logging.UpdateLog(Logging.ServiceName + " : " + "Es el primer empleado.", System.Diagnostics.EventLogEntryType.Error);
                    orgChart.returnCode = -5;
                    orgChart.errorMessage = "Es el primer empleado.";
                }
            }
            catch (Exception ex)
            {
                Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromPreviousEmp : " + "Error al leer el organigrama : " + ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
                orgChart.returnCode = -1;
                orgChart.errorMessage = "No se pudo generar el organigrama, consulte al Administrador.";
            }

            Logging.UpdateLog(Logging.ServiceName + " : " + "ReadOrgChartFromPreviousEmp : " + "Sali", System.Diagnostics.EventLogEntryType.Information);

            return orgChart;
        }

        //Build the orgchart from a specific Employee
        public static OrgChart SaveOrgChart(XElement rootNode)
        {
            OrgChart orgChart = new OrgChart();
            orgChart.orgChartCode = 0;

            Logging.UpdateLog(Logging.ServiceName + " : " + "SaveOrgChart : " + 
                "Entre", System.Diagnostics.EventLogEntryType.Information);

            ConsultasSoapClient proxy = ServiceMan.Get();

            try
            {
                foreach (XElement node in rootNode.Elements())
                {
                    XAttribute att = node.Attribute("codEmp");
                    long codEmp = long.Parse(att.Value);

                    att = node.Attribute("oldCodParent");
                    long oldCodParent = long.Parse(att.Value);

                    att = node.Attribute("newCodParent");
                    long newCodParent = long.Parse(att.Value);

                    proxy.BorrarHijo(codEmp, oldCodParent, BASE);
                    proxy.BorrarPadre(codEmp, oldCodParent, BASE);
                    proxy.AsignarPadre(newCodParent, codEmp, BASE);
                    proxy.AsignarHijo(codEmp, newCodParent, BASE);

                }
            }
            catch (Exception ex)
            {
                Logging.UpdateLog(Logging.ServiceName + " : " + "SaveOrgChart : " + "Error al guardar el organigrama : " + ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
                orgChart.returnCode = -3;
                orgChart.errorMessage = "Error - No se pudieron guardar los cambios realizados en el Organigrama";
            }

            Logging.UpdateLog(Logging.ServiceName + " : " + "SaveOrgChart : " + "Sali", System.Diagnostics.EventLogEntryType.Information);

            return orgChart;
        }

        // Get Employee Data
        private static XElement GetEmployeeData(long codEmp, ConsultasSoapClient proxy, XElement tree)
        {
            //Logging.UpdateLog(Logging.ServiceName + " : " + "GetEmployeeData : " + "Entre", System.Diagnostics.EventLogEntryType.Information);

            tree.SetAttributeValue("empCode", codEmp);

            try
            {
                DataTable empData = proxy.DatosEmp(codEmp, BASE);

                if (empData.Rows.Count > 0)
                {
                    DataRow row = empData.Rows[0];

                    string legajo = row["legajo"].ToString();
                    string nombre = row["nombre"].ToString();
                    string apellido = row["apellido"].ToString();
                    string documento = row["documento"].ToString();
                    string empresa = row["Est1"].ToString();
                    string empresaDesc = row["TipoEst1"].ToString();    //Lisandro Moro
                    string mail = row["mail"].ToString();
                    string interno = row["interno"].ToString();
                    string sucursal = row["Est2"].ToString();
                    string sucursalDesc = row["TipoEst2"].ToString();   //Lisandro Moro
                    string puesto = row["Est3"].ToString();
                    string puestoDesc = row["TipoEst3"].ToString(); //Lisandro Moro
                    string imagefilename = row["Imagen"].ToString();

                    tree.SetAttributeValue("legajo", legajo);
                    tree.SetAttributeValue("nombre", nombre);
                    tree.SetAttributeValue("apellido", apellido);
                    tree.SetAttributeValue("documento", documento);
                    tree.SetAttributeValue("empresa", empresa);
                    tree.SetAttributeValue("empresaDesc", empresaDesc); //Lisandro Moro
                    tree.SetAttributeValue("mail", mail);
                    tree.SetAttributeValue("interno", interno);
                    tree.SetAttributeValue("sucursal", sucursal);
                    tree.SetAttributeValue("sucursalDesc", sucursalDesc);   //Lisandro Moro
                    tree.SetAttributeValue("puesto", puesto);
                    tree.SetAttributeValue("puestoDesc", puestoDesc);   //Lisandro Moro
                    tree.SetAttributeValue("imageFileName", imagefilename);

                    /*
                    string[] itemArr = row.ItemArray.Cast<string>().ToArray();

                    tracer(string.Join(", ", itemArr));*/
                }
            }
            catch (Exception ex)
            {
                Logging.UpdateLog(Logging.ServiceName + " : " + "GetEmployeeData : " + 
                    "Error al obtener los datos del Empleado : " + codEmp + " : " + ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
            }

            //Logging.UpdateLog(Logging.ServiceName + " : " + "GetEmployeeData : " + "Sali", System.Diagnostics.EventLogEntryType.Information);

            return tree;
        }


        private static TextWriterTraceListener ttl = new TextWriterTraceListener(@"d:\yourlog1.log");

        public static void tracer(string pMessage)
        {
            Trace.Listeners.Add(ttl);
            Trace.AutoFlush = true;
            Trace.Indent();
            //Trace.WriteLine("Entering Main");
            Trace.WriteLine(pMessage);
            //Trace.WriteLine("Exiting Main");
            Trace.Unindent();
            Trace.Flush();

        }

        // Build OrgChart from rootEmp
        private static XElement BuildOrgChart(XElement rootEmp, ConsultasSoapClient proxy, int level, int maxLevel)
        {
            try
            {
                DataTable childs = proxy.Hijos(long.Parse(rootEmp.Attribute("empCode").Value), BASE);

                if (childs.Rows.Count > 0)
                {
                    if (level < maxLevel)
                    {
                        foreach (DataRow row in childs.Rows)
                        {
                            long empCode = long.Parse(row["CodEmp"].ToString());
                            XElement empChild = new XElement("child");
                            empChild = GetEmployeeData(empCode, proxy, empChild);
                            empChild = BuildOrgChart(empChild, proxy, level + 1, maxLevel);
                            rootEmp.Add(empChild);
                            
                        }
                    }
                }
            }
            catch(Exception ex){
                
            }
            return rootEmp;
        }
    }
}
