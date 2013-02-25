using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Xml;
using System.Data;
using System.Xml.XPath;

namespace Spreadsheet
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "XMLService" in code, svc and config file together.
    public class XMLService : IXMLService
    {
        BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
        string xmlCode = "code.xml";
        string xmlCost = "cost.xml";

        private XmlDocument getXmlDocumentFromDB(string sheetName)
        {
            XmlDocument xmlDoc = new XmlDocument();

            if (sheetName.Equals(xmlCode))
            {
                XmlElement root = xmlDoc.CreateElement("CODE");
                xmlDoc.AppendChild(root);

                XmlElement root2 = xmlDoc.CreateElement("Service");
                root.AppendChild(root2);
                #region Service
                try{
                    var data = from s in bfAdmin.Services select s;
                    foreach (Service service in data)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        root2.AppendChild(list);

                        XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                        SVCCode.InnerText = service.SVCCode == null ? "" : service.SVCCode.Trim();
                        list.AppendChild(SVCCode);

                        XmlElement SVCName = xmlDoc.CreateElement("SVCName");
                        SVCName.InnerText = service.SVCName == null ? "" : service.SVCName.Trim();
                        list.AppendChild(SVCName);

                        XmlElement SVCDesc = xmlDoc.CreateElement("SVCDesc");
                        SVCDesc.InnerText = service.SVCDesc == null ? "" : service.SVCDesc.Trim();
                        list.AppendChild(SVCDesc);

                        XmlElement ICF_Code = xmlDoc.CreateElement("ICF_Code");
                        ICF_Code.InnerText = service.ICF_Code == null ? "" : service.ICF_Code.Trim();
                        list.AppendChild(ICF_Code);

                        XmlElement HostCode = xmlDoc.CreateElement("ProviderCode");
                        HostCode.InnerText = service.ProviderCode == null ? "" : service.ProviderCode.Trim();
                        list.AppendChild(HostCode);

                        XmlElement StaffRole = xmlDoc.CreateElement("StaffRole");
                        StaffRole.InnerText = service.StaffRole == null ? "" : service.StaffRole.Trim();
                        list.AppendChild(StaffRole);

                        XmlElement SVCType = xmlDoc.CreateElement("SVCType");
                        SVCType.InnerText = service.SVCType == null ? "" : service.SVCType.Trim();
                        list.AppendChild(SVCType);

                        XmlElement SVCObjective = xmlDoc.CreateElement("SVCObjective");
                        SVCObjective.InnerText = service.SVCObjective == null ? "" : service.SVCObjective.Trim();
                        list.AppendChild(SVCObjective);

                        XmlElement SVCSupport = xmlDoc.CreateElement("SVCSupport");
                        SVCSupport.InnerText = service.SVCSupport == null ? "" : service.SVCSupport.Trim();
                        list.AppendChild(SVCSupport);

                        XmlElement SVCCoverage = xmlDoc.CreateElement("SVCCoverage");
                        SVCCoverage.InnerText = service.SVCCoverage == null ? "" : service.SVCCoverage.Trim();
                        list.AppendChild(SVCCoverage);

                        XmlElement SVCStart = xmlDoc.CreateElement("SVCStart");
                        SVCStart.InnerText = service.SVCStart == null ? "" : service.SVCStart.Trim();
                        list.AppendChild(SVCStart);

                        XmlElement SVCEnd = xmlDoc.CreateElement("SVCEnd");
                        SVCEnd.InnerText = service.SVCEnd == null ? "" : service.SVCEnd.Trim();
                        list.AppendChild(SVCEnd);

                        XmlElement ChildType = xmlDoc.CreateElement("ChildType");
                        ChildType.InnerText = service.ChildType == null ? "" : service.ChildType.Trim();
                        list.AppendChild(ChildType);

                        XmlElement ProblemToSolve = xmlDoc.CreateElement("ProblemToSolve");
                        ProblemToSolve.InnerText = service.ProblemToSolve == null ? "" : service.ProblemToSolve.Trim();
                        list.AppendChild(ProblemToSolve);

                        XmlElement OrgAssignedCode = xmlDoc.CreateElement("OrgAssignedCode");
                        OrgAssignedCode.InnerText = service.OrgAssignedCode == null ? "" : service.OrgAssignedCode.Trim();
                        list.AppendChild(OrgAssignedCode);
                    }
                }
                catch(Exception){ }
                 #endregion
                
                XmlElement rootAct = xmlDoc.CreateElement("Activity");
                root.AppendChild(rootAct);
                #region Activity
                var dataAct = from s in bfAdmin.Activities select s;
                        foreach(Activity activity in dataAct)
                        {
                                XmlElement list = xmlDoc.CreateElement("DATA");
                                rootAct.AppendChild(list);

                                XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                                ACTCode.InnerText = activity.ACTCode == null ? "" : activity.ACTCode.Trim();
                                list.AppendChild(ACTCode);

                                XmlElement ACTDes = xmlDoc.CreateElement("ACTDest");
                                ACTDes.InnerText = activity.ACTDesc == null ? "" : activity.ACTDesc.Trim();
                                list.AppendChild(ACTDes);

                                XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                                SVCCode.InnerText = activity.SVCCode == null ? "" : activity.SVCCode.Trim();
                                list.AppendChild(SVCCode);

                                XmlElement ICF_Code = xmlDoc.CreateElement("ICF_Code");
                                ICF_Code.InnerText = activity.ICF_Code == null ? "" : activity.ICF_Code.Trim();
                                list.AppendChild(ICF_Code);
                         }
                        
                    #endregion

                XmlElement rootSAct = xmlDoc.CreateElement("SubActivity");
                root.AppendChild(rootSAct);
                #region SubActivity
                try
                {
                    var dataSAct = from s in bfAdmin.SubActivities select s;
                    foreach (SubActivity sactivity in dataSAct)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        rootSAct.AppendChild(list);

                        XmlElement SACTCode = xmlDoc.CreateElement("SACTCode");
                        SACTCode.InnerText = sactivity.SACTCode == null ? "" : sactivity.SACTCode.Trim();
                        list.AppendChild(SACTCode);

                        XmlElement SACTDes = xmlDoc.CreateElement("SACTDest");
                        SACTDes.InnerText = sactivity.SACTDesc == null ? "" : sactivity.SACTDesc.Trim();
                        list.AppendChild(SACTDes);

                        XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                        ACTCode.InnerText = sactivity.ACTCode == null ? "" : sactivity.ACTCode.Trim();
                        list.AppendChild(ACTCode);

                        XmlElement ICF_Code = xmlDoc.CreateElement("ICF_Code");
                        ICF_Code.InnerText = sactivity.ICF_Code == null ? "" : sactivity.ICF_Code.Trim();
                        list.AppendChild(ICF_Code);
                    }
                }
                catch (Exception) { }
                #endregion

                XmlElement rootMat = xmlDoc.CreateElement("Material");
                root.AppendChild(rootMat);
                #region Material
                try
                {
                    var dataMat = from s in bfAdmin.Materials select s;
                    foreach (Material material in dataMat)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        rootMat.AppendChild(list);

                        XmlElement MaterialCode = xmlDoc.CreateElement("MaterialCode");
                        MaterialCode.InnerText = material.MaterialCode == null ? "" : material.MaterialCode.Trim();
                        list.AppendChild(MaterialCode);

                        XmlElement MaterialDesc = xmlDoc.CreateElement("MaterialDesc");
                        MaterialDesc.InnerText = material.MaterialDesc == null ? "" : material.MaterialDesc.Trim();
                        list.AppendChild(MaterialDesc);

                        XmlElement Unit = xmlDoc.CreateElement("Unit");
                        Unit.InnerText = material.Unit == null ? "" : material.Unit.Trim();
                        list.AppendChild(Unit);

                        XmlElement EstimatedPrice = xmlDoc.CreateElement("EstimatedPrice");
                        EstimatedPrice.InnerText = material.EstimatedPrice == null ? "" : material.EstimatedPrice.Trim();
                        list.AppendChild(EstimatedPrice);

                        XmlElement RealPrice = xmlDoc.CreateElement("RealPrice");
                        RealPrice.InnerText = material.RealPrice == null ? "" : material.RealPrice.Trim();
                        list.AppendChild(RealPrice);

                        XmlElement SVCCode = xmlDoc.CreateElement("ACTCode");
                        SVCCode.InnerText = material.ACTCode == null ? "" : material.ACTCode.Trim();
                        list.AppendChild(SVCCode);

                        XmlElement Note = xmlDoc.CreateElement("Note");
                        Note.InnerText = material.Note == null ? "" : material.Note.Trim();
                        list.AppendChild(Note);
                    }
                }
                catch(Exception){ }
                #endregion 
            }
            else if (sheetName.Equals(xmlCost))
            {
                XmlElement root = xmlDoc.CreateElement("COST");
                xmlDoc.AppendChild(root);
        
                XmlElement rootAnno = xmlDoc.CreateElement("Annotation");
                root.AppendChild(rootAnno);
                #region Annotation
                try
                {
                    var dataAnno = from s in bfAdmin.Annotations select s;
                    foreach (Annotation annotation in dataAnno)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        rootAnno.AppendChild(list);

                        XmlElement AID = xmlDoc.CreateElement("AID");
                        AID.InnerText = annotation.AID == null ? "" : annotation.AID.Trim();
                        list.AppendChild(AID);

                        XmlElement AText = xmlDoc.CreateElement("AText");
                        AText.InnerText = annotation.AText == null ? "" : annotation.AText.Trim();
                        list.AppendChild(AText);

                        XmlElement AnnotationID = xmlDoc.CreateElement("AnnotationID");
                        AnnotationID.InnerText = annotation.AnnotationID == null ? "" : annotation.AnnotationID.Trim();
                        list.AppendChild(AnnotationID);

                        XmlElement Reference = xmlDoc.CreateElement("Reference");
                        Reference.InnerText = annotation.Reference == null ? "" : annotation.Reference.Trim();
                        list.AppendChild(Reference);
                    }
                }
                catch (Exception) { }
                #endregion

                XmlElement rootACost = xmlDoc.CreateElement("ActivityCost");
                root.AppendChild(rootACost);
                #region ActivityCost
                try
                {
                    var dataACost = from s in bfAdmin.ActivityCosts select s;
                    foreach (ActivityCost activityCost in dataACost)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        rootACost.AppendChild(list);

                        XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                        ACTCode.InnerText = activityCost.ACTCode == null ? "" : activityCost.ACTCode.Trim();
                        list.AppendChild(ACTCode);

                        XmlElement Unit = xmlDoc.CreateElement("Unit");
                        Unit.InnerText = activityCost.Unit == null ? "" : activityCost.Unit.Trim();
                        list.AppendChild(Unit);

                        XmlElement LabourCost = xmlDoc.CreateElement("LabourCost");
                        LabourCost.InnerText = activityCost.LabourCost == null ? "" : activityCost.LabourCost.Trim();
                        list.AppendChild(LabourCost);

                        XmlElement MaterialCost = xmlDoc.CreateElement("MaterialCost");
                        MaterialCost.InnerText = activityCost.MaterialCost == null ? "" : activityCost.MaterialCost.Trim();
                        list.AppendChild(MaterialCost);

                        XmlElement CC_Equipment = xmlDoc.CreateElement("CC_Equipment");
                        CC_Equipment.InnerText = activityCost.CC_Equipment == null ? "" : activityCost.CC_Equipment.Trim();
                        list.AppendChild(CC_Equipment);

                        XmlElement CC_Building = xmlDoc.CreateElement("CC_Building");
                        CC_Building.InnerText = activityCost.CC_Building == null ? "" : activityCost.CC_Building.Trim();
                        list.AppendChild(CC_Building);

                        XmlElement IndirectCost = xmlDoc.CreateElement("IndirectCost");
                        IndirectCost.InnerText = activityCost.IndirectCost == null ? "" : activityCost.IndirectCost.Trim();
                        list.AppendChild(IndirectCost);

                        XmlElement ProposedCost = xmlDoc.CreateElement("ProposedCost");
                        ProposedCost.InnerText = activityCost.ProposedCost == null ? "" : activityCost.ProposedCost.Trim();
                        list.AppendChild(ProposedCost);

                        XmlElement CurrentCost = xmlDoc.CreateElement("CurrentCost");
                        CurrentCost.InnerText = activityCost.CurrentCost == null ? "" : activityCost.CurrentCost.Trim();
                        list.AppendChild(CurrentCost);

                        XmlElement UnitCost = xmlDoc.CreateElement("UnitCost");
                        UnitCost.InnerText = activityCost.UnitCost == null ? "" : activityCost.UnitCost.Trim();
                        list.AppendChild(UnitCost);

                        XmlElement ReferencedCostOrg = xmlDoc.CreateElement("ReferencedCostOrg");
                        ReferencedCostOrg.InnerText = activityCost.ReferencedCostOrg == null ? "" : activityCost.ReferencedCostOrg.Trim();
                        list.AppendChild(ReferencedCostOrg);

                        XmlElement TimsStamp = xmlDoc.CreateElement("TimsStamp");
                        TimsStamp.InnerText = activityCost.TimsStamp == null ? "" : activityCost.TimsStamp.ToString();
                        list.AppendChild(TimsStamp);

                        XmlElement AID = xmlDoc.CreateElement("AID");
                        AID.InnerText = activityCost.AID == null ? "" : activityCost.AID.Trim();
                        list.AppendChild(AID);

                    }
                }
                catch (Exception) { }
                #endregion
                            
            }
            return xmlDoc;
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
            bfAdmin = new BenefitAdminDataContext();
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

                        XmlElement ICF_Code = xmlDoc.CreateElement("ICF_Code");
                        ICF_Code.InnerXml = s.ICF_Code.Trim();
                        act.AppendChild(ICF_Code);
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

                        XmlElement ICF_Code = xmlDoc.CreateElement("ICF_Code");
                        ICF_Code.InnerXml = s.ICF_Code.Trim();
                        sact.AppendChild(ICF_Code);
                    }
                }
                #endregion
                else if (fk.Equals("Material") && pk.Equals("Service"))
                #region
                {
                    var data = from m in bfAdmin.Materials where stringPK.Trim().Equals(m.ACTCode.Trim()) select m;
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
                        SVCCode.InnerXml = s.ACTCode.Trim();
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
         * input : name of organization and name of sheet
         *       : name of organization and string "All"
         *       : name of organization and string "List"
         * output : XmlElement specific by input
         * contract string string -> XmlElement
         * */
        public XmlElement getXMLSheet(string group, string sheet)
        {
            bfAdmin = new BenefitAdminDataContext();
            XmlDocument xmlDoc = new XmlDocument();
            //init filepath and read xml document
            #region
            if (group.ToUpper().Equals("CODE"))
            {
                if (sheet.ToUpper().Equals("ALL"))
                {
                    return getXmlDocumentFromDB(xmlCode).DocumentElement;
                }
                else
                {
                    xmlDoc = getXmlDocumentFromDB(xmlCode);
                }
            }
            else if (group.ToUpper().Equals("COST"))
            {
                if (sheet.ToUpper().Equals("ALL"))
                {
                    return getXmlDocumentFromDB(xmlCost).DocumentElement;
                }
                else
                {
                    xmlDoc = getXmlDocumentFromDB(xmlCost);
                }
            }
            #endregion

            XmlNodeList nodes = xmlDoc.GetElementsByTagName(group.ToUpper());
            if (!sheet.ToUpper().Equals("LIST"))
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
                if (group.ToUpper().Equals("COST"))
                {
                    XmlDocument doc = new XmlDocument();
                    XmlElement root = doc.CreateElement("Lists");
                    doc.AppendChild(root);
                    XmlElement list = doc.CreateElement("SheetName");
                    list.InnerText = "Annotation";
                    root.AppendChild(list);
                    list = doc.CreateElement("SheetName");
                    list.InnerText = "ActivityCost";
                    root.AppendChild(list);
                    return doc.DocumentElement;
                }
                else if (group.ToUpper().Equals("CODE"))
                {
                    XmlDocument doc = new XmlDocument();
                    XmlElement root = doc.CreateElement("Lists");
                    doc.AppendChild(root);
                    XmlElement list = doc.CreateElement("SheetName");
                    list.InnerText = "Service";
                    root.AppendChild(list);
                    list = doc.CreateElement("SheetName");
                    list.InnerText = "Activity";
                    root.AppendChild(list);
                    list = doc.CreateElement("SheetName");
                    list.InnerText = "SubActivity";
                    root.AppendChild(list);
                    list = doc.CreateElement("SheetName");
                    list.InnerText = "Material";
                    root.AppendChild(list);
                    return doc.DocumentElement;
                }
            }
            return null;
        }
    }
}
