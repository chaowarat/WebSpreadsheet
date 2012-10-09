using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Xml;
using System.Data;

namespace Spreadsheet
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "XMLService" in code, svc and config file together.
    public class XMLService : IXMLService
    {
        BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();

        private XmlElement getCodeXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            string pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\code.xml";
            xmlDoc.Load(pathSheets);

            return xmlDoc.DocumentElement;
        }

        private XmlElement getProviderXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            string pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\provider.xml";
            xmlDoc.Load(pathSheets);

            return xmlDoc.DocumentElement;
        }

        private XmlElement getCostXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            string pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\cost.xml";
            xmlDoc.Load(pathSheets);

            return xmlDoc.DocumentElement;
        }

        private List<string> queryPK(string table)
        {
            if (table.Equals("Service"))
            {
                var data = from m in bfAdmin.Services select m.SVCCode;
                return data.ToList<string>();
            }
            else if (table.Equals("Activity"))
            {
                var data = from m in bfAdmin.Activities select m.ACTCode;
                return data.ToList<string>();
            }
            else if (table.Equals("SubActivity"))
            {
                var data = from m in bfAdmin.SubActivities select m.SACTCode;
                return data.ToList<string>();
            }
            else if (table.Equals("Material"))
            {
                var data = from m in bfAdmin.Materials select m.MaterialCode;
                return data.ToList<string>();
            }
            return null;
        }

        public XmlElement getXMLSheet(string select, string pk, string fk)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement dataElm = xmlDoc.CreateElement("DATA");
            xmlDoc.AppendChild(dataElm);

            List<string> dataPK = this.queryPK(pk);
            if (!select.Equals("All"))
            {
                string tmp = "";
                foreach (string a in dataPK)
                {
                    if (a.Trim().Equals(select))
                    {
                        tmp = a.Trim();
                    }
                }
                dataPK.Clear();
                dataPK.Add(tmp);
            }

            foreach (string stringPK in dataPK)
            {
                XmlElement root = xmlDoc.CreateElement(pk);
                root.SetAttribute("code", stringPK.Trim());
                dataElm.AppendChild(root);

                if (fk.Equals("Activity") && pk.Equals("Service"))
                #region
                {
                    var data = from m in bfAdmin.Activities where stringPK.Trim().Equals(m.SVCCode.Trim()) select m;
                    foreach (var s in data)
                    {
                        XmlElement act = xmlDoc.CreateElement("Activity");
                        root.AppendChild(act);

                        XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                        ACTCode.InnerXml = s.ACTCode.Trim();
                        act.AppendChild(ACTCode);

                        XmlElement ACTDesc = xmlDoc.CreateElement("ACTDesc");
                        ACTDesc.InnerXml = s.ACTDesc.Trim();
                        act.AppendChild(ACTDesc);

                        XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                        SVCCode.InnerXml = s.SVCCode.Trim();
                        act.AppendChild(SVCCode);
                    }
                }
                #endregion
                else if (fk.Equals("SubActivity") && pk.Equals("Activity"))
                #region
                {
                    var data = from m in bfAdmin.SubActivities where stringPK.Trim().Equals(m.ACTCode.Trim()) select m;
                    foreach (var s in data)
                    {
                        XmlElement sact = xmlDoc.CreateElement("SubActivity");
                        root.AppendChild(sact);

                        XmlElement SACTCode = xmlDoc.CreateElement("SACTCode");
                        SACTCode.InnerXml = s.SACTCode.Trim();
                        sact.AppendChild(SACTCode);

                        XmlElement SACTDesc = xmlDoc.CreateElement("SACTDesc");
                        SACTDesc.InnerXml = s.SACTDesc.Trim();
                        sact.AppendChild(SACTDesc);

                        XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                        ACTCode.InnerXml = s.ACTCode.Trim();
                        sact.AppendChild(ACTCode);
                    }
                }
                #endregion
                else if (fk.Equals("Material") && pk.Equals("Service"))
                #region
                {
                    var data = from m in bfAdmin.Materials where stringPK.Trim().Equals(m.SVCCode.Trim()) select m;
                    foreach (var s in data)
                    {
                        XmlElement mat = xmlDoc.CreateElement("Material");
                        root.AppendChild(mat);

                        XmlElement MaterialCode = xmlDoc.CreateElement("MaterialCode");
                        MaterialCode.InnerXml = s.MaterialCode.Trim();
                        mat.AppendChild(MaterialCode);

                        XmlElement MaterialDesc = xmlDoc.CreateElement("MaterialDesc");
                        MaterialDesc.InnerXml = s.MaterialDesc.Trim();
                        mat.AppendChild(MaterialDesc);

                        XmlElement Unit = xmlDoc.CreateElement("Unit");
                        Unit.InnerXml = s.Unit.Trim();
                        mat.AppendChild(Unit);

                        XmlElement EstimatedPrice = xmlDoc.CreateElement("EstimatedPrice");
                        EstimatedPrice.InnerXml = s.EstimatedPrice.Trim();
                        mat.AppendChild(EstimatedPrice);

                        XmlElement RealPrice = xmlDoc.CreateElement("RealPrice");
                        RealPrice.InnerXml = s.RealPrice.Trim();
                        mat.AppendChild(RealPrice);

                        XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                        SVCCode.InnerXml = s.SVCCode.Trim();
                        mat.AppendChild(SVCCode);

                        XmlElement Note = xmlDoc.CreateElement("Note");
                        Note.InnerXml = s.Note.Trim();
                        mat.AppendChild(Note);
                    }
                }
                #endregion
            }

            return xmlDoc.DocumentElement;
        }

        /* getXMLSheet : returen specific XmlElement of sheet
         * input : name of organization and name of sheet(can use "All" or "List" instead sheet name)
         *       : name of organization and name of sheet
         *       : name of organization and string "All"
         *       : name of organization and string "List"
         * output : XmlElement specific by input
         * contract string string -> XmlElement
         * */
        public XmlElement getXMLSheet(string group, string sheet)
        {
            XmlDocument xmlDoc = new XmlDocument();
            string pathSheets = "";
            //init filepath and read xml document
            #region
            if (group.Equals("Code"))
            {
                if (sheet.Equals("All"))
                {
                    return getCodeXML();
                }
                else
                {
                    pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\code.xml";
                    xmlDoc.Load(pathSheets);
                }
            }
            else if (group.Equals("Cost"))
            {
                if (sheet.Equals("All"))
                {
                    return getCostXML();
                }
                else
                {
                    pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\cost.xml";
                    xmlDoc.Load(pathSheets);
                }
            }
            else if (group.Equals("Provider"))
            {
                if (sheet.Equals("All"))
                {
                    return getProviderXML();
                }
                else
                {
                    pathSheets = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"sheets\provider.xml";
                    xmlDoc.Load(pathSheets);
                }
            }
            #endregion

            if (group.ToUpper().Equals("COST"))
            {
                XmlDocument doc = new XmlDocument();
                XmlElement root = doc.CreateElement("Lists");
                doc.AppendChild(root);
                XmlElement list = doc.CreateElement("SheetName");
                list.InnerText = "ActivityCost";
                root.AppendChild(list);
                return doc.DocumentElement;
            }
            XmlNodeList nodes = xmlDoc.GetElementsByTagName(group.ToUpper());
            if (!sheet.Equals("List"))
            {
                foreach (XmlNode node in nodes.Item(0).ChildNodes)
                {
                    if (node.Name.Equals(sheet))
                    {
                        return (XmlElement)node;
                    }
                }
                return null;
            }
            else
            {
                XmlDocument doc = new XmlDocument();
                XmlElement root = doc.CreateElement("Lists");
                doc.AppendChild(root);
                foreach (XmlNode node in nodes.Item(0).ChildNodes)
                {
                    XmlElement list = doc.CreateElement("SheetName");
                    list.InnerText = node.Name;
                    root.AppendChild(list);
                }
                return doc.DocumentElement;
            }
        }
    }
}
