namespace RhPro.Oodd.Web.Services
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;
    using System.ServiceModel.DomainServices.Hosting;
    using System.ServiceModel.DomainServices.Server;

    using RhPro.Oodd.Web.DataModel;
    using RhPro.Oodd.Web.OrgDao;
    using System.Xml.Linq;

    // TODO: Create methods containing your application logic.
    [EnableClientAccess()]
    public class OrgDomainService : DomainService
    {

        [Invoke]
        public OrgChart ReadOrg(long legajo, int maxLevel)
        {
            OrgChart orgChart = OrgDaoHandler.ReadOrgChart(legajo, maxLevel);
            return orgChart;
            //return new OrgChart();
        }

        [Invoke]
        public OrgChart ReadOrgFromNextEmp(long legajo, int maxLevel)
        {
            OrgChart orgChart = OrgDaoHandler.ReadOrgChartFromNextEmp(legajo, maxLevel);
            return orgChart;
            //return new OrgChart();

        }

        [Invoke]
        public OrgChart ReadOrgFromPreviousEmp(long legajo, int maxLevel)
        {
            OrgChart orgChart = OrgDaoHandler.ReadOrgChartFromPreviousEmp(legajo, maxLevel);
            return orgChart;
            //return new OrgChart();

        }

        [Invoke]
        //public OrgChart SaveOrg(XElement nodes)
        public OrgChart SaveOrg(XElement nodes)
        {
            OrgChart orgChart = OrgDaoHandler.SaveOrgChart(nodes);
            return orgChart;
            //return new OrgChart();

        }

        public IEnumerable<OrgChart> GetMyOrgChart()
        {
            throw new NotImplementedException();
        }
    }
}


