﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO;
using System.Xml.Linq;
using System.Web;
using System.Xml.XPath;
using System.Text;

namespace Spreadsheet
{
    public class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding { get { return Encoding.UTF8; } }
    }

    public class HtmlManager
    {
        static BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
        static private string pathSheets = System.Web.HttpContext.Current.Server.MapPath("~/sheets");
        static private string pathTemporary = pathSheets + @"/temporary.html";
        static int countCol = 2;
        static string border = "1px";
        static string id = "jSheet_0_0";
        static string classsheet = "jSheet ui-widget-content";
        static int count = 0;
        static XElement TBody, td, countrow;

        public static void createHtml(string fileName)
        {
            initFile();
            FileStream fs = new FileStream(pathTemporary, FileMode.Append);
            StreamWriter h = new StreamWriter(fs, System.Text.Encoding.UTF8);

            if (System.IO.File.Exists(pathSheets + @"/" + fileName))
            {
                StreamReader read = new StreamReader(pathSheets + @"/" + fileName);
                h.Write(read.ReadToEnd());
                h.Close();
                fs.Close();
                read.Close();
            }
            else
            {
                if (fileName.Equals("provider.html"))
                {
                    h.Write(createMinistryHtml());
                    h.Write(createOrganizationHtml());
                    h.Write(createProviderHtml());
                }
                else if (fileName.Equals("cost.html"))
                {
                    h.Write(createActivityCostFromDB());
                }
                else if (fileName.Equals("code.html"))
                {
                    h.Write(createServiceFromDB());
                    h.Write(createActivityFromDB());
                    h.Write(createSubActivityFromDB());
                    h.Write(createMaterialFromDB());
                }

                h.Close();
                fs.Close();
            }
        }

        public static void saveHtml(string data, string fileName)
        {
            FileStream fs = new FileStream(pathSheets + @"/" + fileName, FileMode.Create);
            StreamWriter writer = new StreamWriter(fs, System.Text.Encoding.UTF8);
            XmlDocument xmlDoc = new XmlDocument();
            StringReader rdr = new StringReader(data);
            XPathDocument document = new XPathDocument(rdr);
            Utf8StringWriter sw = new Utf8StringWriter();

            if (fileName.Equals("cost.xml"))
            {
                #region
                XPathNavigator navigator = document.CreateNavigator();
                XPathNodeIterator allDocs = navigator.Select("DOCUMENTS/DOCUMENT/DATA");
                XPathNodeIterator row = navigator.Select("DOCUMENTS/DOCUMENT/METADATA/ROWS");
                row.MoveNext();
                int countRow = Convert.ToInt32(row.Current.Value);
                xmlDoc = new XmlDocument();
                XmlElement root = xmlDoc.CreateElement("ActivityCost");
                xmlDoc.AppendChild(root);

                while (allDocs.MoveNext())
                {   
                    for (int i = 1; i < countRow; i++)
                    {
                        XmlElement list = xmlDoc.CreateElement("DATA");
                        root.AppendChild(list);

                        XPathNodeIterator getCol0 = allDocs.Current.Select("R" + i + "/C0");
                        getCol0.MoveNext();
                        XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                        ACTCode.InnerText = getCol0.Current.Value.Trim();
                        list.AppendChild(ACTCode);

                        XPathNodeIterator getCol1 = allDocs.Current.Select("R" + i + "/C1");
                        getCol1.MoveNext();
                        XmlElement ACTName = xmlDoc.CreateElement("ACTName");
                        ACTName.InnerText = getCol1.Current.Value.Trim();
                        list.AppendChild(ACTName);

                        XPathNodeIterator getCol2 = allDocs.Current.Select("R" + i + "/C2");
                        getCol2.MoveNext();
                        XmlElement Unit = xmlDoc.CreateElement("Unit");
                        Unit.InnerText = getCol2.Current.Value.Trim();
                        list.AppendChild(Unit);

                        XPathNodeIterator getCol3 = allDocs.Current.Select("R" + i + "/C3");
                        getCol3.MoveNext();
                        XmlElement LabourCost = xmlDoc.CreateElement("LabourCost");
                        LabourCost.InnerText = getCol3.Current.Value.Trim();
                        list.AppendChild(LabourCost);

                        XPathNodeIterator getCol4 = allDocs.Current.Select("R" + i + "/C4");
                        getCol4.MoveNext();
                        XmlElement MaterialCost = xmlDoc.CreateElement("MaterialCost");
                        MaterialCost.InnerText = getCol4.Current.Value.Trim();
                        list.AppendChild(MaterialCost);

                        XPathNodeIterator getCol5 = allDocs.Current.Select("R" + i + "/C5");
                        getCol5.MoveNext();
                        XmlElement CC_Equipment = xmlDoc.CreateElement("CC_Equipment");
                        CC_Equipment.InnerText = getCol5.Current.Value.Trim();
                        list.AppendChild(CC_Equipment);

                        XPathNodeIterator getCol6 = allDocs.Current.Select("R" + i + "/C6");
                        getCol6.MoveNext();
                        XmlElement CC_Building = xmlDoc.CreateElement("CC_Building");
                        CC_Building.InnerText = getCol6.Current.Value.Trim();
                        list.AppendChild(CC_Building);

                        XPathNodeIterator getCol7 = allDocs.Current.Select("R" + i + "/C7");
                        getCol7.MoveNext();
                        XmlElement IndirectCost = xmlDoc.CreateElement("IndirectCost");
                        IndirectCost.InnerText = getCol7.Current.Value.Trim();
                        list.AppendChild(IndirectCost);

                        XPathNodeIterator getCol8 = allDocs.Current.Select("R" + i + "/C8");
                        getCol8.MoveNext();
                        XmlElement ProposedCost = xmlDoc.CreateElement("ProposedCost");
                        ProposedCost.InnerText = getCol8.Current.Value.Trim();
                        list.AppendChild(ProposedCost);

                        XPathNodeIterator getCol9 = allDocs.Current.Select("R" + i + "/C9");
                        getCol9.MoveNext();
                        XmlElement CurrentCost = xmlDoc.CreateElement("CurrentCost");
                        CurrentCost.InnerText = getCol9.Current.Value.Trim();
                        list.AppendChild(CurrentCost);

                        XPathNodeIterator getCol10 = allDocs.Current.Select("R" + i + "/C10");
                        getCol10.MoveNext();
                        XmlElement UnitCost = xmlDoc.CreateElement("UnitCost");
                        UnitCost.InnerText = getCol10.Current.Value.Trim();
                        list.AppendChild(UnitCost);

                        XPathNodeIterator getCol11 = allDocs.Current.Select("R" + i + "/C11");
                        getCol11.MoveNext();
                        XmlElement ReferencedCostOrg = xmlDoc.CreateElement("ReferencedCostOrg");
                        ReferencedCostOrg.InnerText = getCol11.Current.Value.Trim();
                        list.AppendChild(ReferencedCostOrg);
                    }
                }
                #endregion
            }
            else if (fileName.Equals("code.xml"))
            {
                XPathNavigator navigator = document.CreateNavigator();
                XPathNodeIterator allDocs = navigator.Select("DOCUMENTS/DOCUMENT");
                XmlElement root = xmlDoc.CreateElement("CODE");
                xmlDoc.AppendChild(root);
                while (allDocs.MoveNext())
                {
                    string table = allDocs.Current.GetAttribute("title", "");
                    if (table.Equals("Service"))
                    {
                        #region
                        XPathNodeIterator row = navigator.Select("DOCUMENTS/DOCUMENT[@title='Service']/METADATA/ROWS");
                        row.MoveNext();
                        int countRow = Convert.ToInt32(row.Current.Value);
                        //xmlDoc = new XmlDocument();
                        XmlElement root2 = xmlDoc.CreateElement("Service");
                        root.AppendChild(root2);
                        XPathNodeIterator inDocs = navigator.Select("DOCUMENTS/DOCUMENT[@title='Service']/DATA");
                        while (inDocs.MoveNext())
                        {
                            for (int i = 1; i < countRow; i++)
                            {
                                XmlElement list = xmlDoc.CreateElement("DATA");
                                root2.AppendChild(list);

                                XPathNodeIterator getCol0 = inDocs.Current.Select("R" + i + "/C0");
                                getCol0.MoveNext();
                                XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                                SVCCode.InnerText = getCol0.Current.Value.Trim();
                                list.AppendChild(SVCCode);

                                XPathNodeIterator getCol1 = inDocs.Current.Select("R" + i + "/C1");
                                getCol1.MoveNext();
                                XmlElement SVCName = xmlDoc.CreateElement("SVCName");
                                SVCName.InnerText = getCol1.Current.Value.Trim();
                                list.AppendChild(SVCName);

                                XPathNodeIterator getCol2 = inDocs.Current.Select("R" + i + "/C2");
                                getCol2.MoveNext();
                                XmlElement SVCDesc = xmlDoc.CreateElement("SVCDesc");
                                SVCDesc.InnerText = getCol2.Current.Value.Trim();
                                list.AppendChild(SVCDesc);

                                XPathNodeIterator getCol3 = inDocs.Current.Select("R" + i + "/C3");
                                getCol3.MoveNext();
                                XmlElement HostCode = xmlDoc.CreateElement("HostCode");
                                HostCode.InnerText = getCol3.Current.Value.Trim();
                                list.AppendChild(HostCode);

                                XPathNodeIterator getCol4 = inDocs.Current.Select("R" + i + "/C4");
                                getCol4.MoveNext();
                                XmlElement StaffRole = xmlDoc.CreateElement("StaffRole");
                                StaffRole.InnerText = getCol4.Current.Value.Trim();
                                list.AppendChild(StaffRole);

                                XPathNodeIterator getCol5 = inDocs.Current.Select("R" + i + "/C5");
                                getCol5.MoveNext();
                                XmlElement SVCType = xmlDoc.CreateElement("SVCType");
                                SVCType.InnerText = getCol5.Current.Value.Trim();
                                list.AppendChild(SVCType);

                                XPathNodeIterator getCol6 = inDocs.Current.Select("R" + i + "/C6");
                                getCol6.MoveNext();
                                XmlElement SVCObjective = xmlDoc.CreateElement("SVCObjective");
                                SVCObjective.InnerText = getCol6.Current.Value.Trim();
                                list.AppendChild(SVCObjective);

                                XPathNodeIterator getCol7 = inDocs.Current.Select("R" + i + "/C7");
                                getCol7.MoveNext();
                                XmlElement SVCSupport = xmlDoc.CreateElement("SVCSupport");
                                SVCSupport.InnerText = getCol7.Current.Value.Trim();
                                list.AppendChild(SVCSupport);

                                XPathNodeIterator getCol8 = inDocs.Current.Select("R" + i + "/C8");
                                getCol8.MoveNext();
                                XmlElement SVCCoverage = xmlDoc.CreateElement("SVCCoverage");
                                SVCCoverage.InnerText = getCol8.Current.Value.Trim();
                                list.AppendChild(SVCCoverage);

                                XPathNodeIterator getCol9 = inDocs.Current.Select("R" + i + "/C9");
                                getCol9.MoveNext();
                                XmlElement SVCStart = xmlDoc.CreateElement("SVCStart");
                                SVCStart.InnerText = getCol9.Current.Value.Trim();
                                list.AppendChild(SVCStart);

                                XPathNodeIterator getCol10 = inDocs.Current.Select("R" + i + "/C10");
                                getCol10.MoveNext();
                                XmlElement SVCEnd = xmlDoc.CreateElement("SVCEnd");
                                SVCEnd.InnerText = getCol10.Current.Value.Trim();
                                list.AppendChild(SVCEnd);
                            }
                        }
                        #endregion
                    }
                    else if (table.Equals("Activity"))
                    {
                        #region
                        XPathNodeIterator row = navigator.Select("DOCUMENTS/DOCUMENT[@title='Activity']/METADATA/ROWS");
                        row.MoveNext();
                        int countRow = Convert.ToInt32(row.Current.Value);
                        XmlElement root2 = xmlDoc.CreateElement("Activity");
                        root.AppendChild(root2);
                        XPathNodeIterator inDocs = navigator.Select("DOCUMENTS/DOCUMENT[@title='Activity']/DATA");
                        while (inDocs.MoveNext())
                        {
                            for (int i = 1; i < countRow; i++)
                            {
                                XmlElement list = xmlDoc.CreateElement("DATA");
                                root2.AppendChild(list);

                                XPathNodeIterator getCol0 = inDocs.Current.Select("R" + i + "/C0");
                                getCol0.MoveNext();
                                XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                                ACTCode.InnerText = getCol0.Current.Value.Trim();
                                list.AppendChild(ACTCode);

                                XPathNodeIterator getCol1 = inDocs.Current.Select("R" + i + "/C1");
                                getCol1.MoveNext();
                                XmlElement ACTDes = xmlDoc.CreateElement("ACTDest");
                                ACTDes.InnerText = getCol1.Current.Value.Trim();
                                list.AppendChild(ACTDes);

                                XPathNodeIterator getCol2 = inDocs.Current.Select("R" + i + "/C2");
                                getCol2.MoveNext();
                                XmlElement SVCCode = xmlDoc.CreateElement("SVCCode");
                                SVCCode.InnerText = getCol2.Current.Value.Trim();
                                list.AppendChild(SVCCode);
                            }
                        }
                        #endregion
                    }
                    else if (table.Equals("SubActivity"))
                    {
                        #region
                        XPathNodeIterator row = navigator.Select("DOCUMENTS/DOCUMENT[@title='SubActivity']/METADATA/ROWS");
                        row.MoveNext();
                        int countRow = Convert.ToInt32(row.Current.Value);
                        XmlElement root2 = xmlDoc.CreateElement("SubActivity");
                        root.AppendChild(root2);
                        XPathNodeIterator inDocs = navigator.Select("DOCUMENTS/DOCUMENT[@title='SubActivity']/DATA");
                        while (inDocs.MoveNext())
                        {
                            for (int i = 1; i < countRow; i++)
                            {
                                XmlElement list = xmlDoc.CreateElement("DATA");
                                root2.AppendChild(list);

                                XPathNodeIterator getCol0 = inDocs.Current.Select("R" + i + "/C0");
                                getCol0.MoveNext();
                                XmlElement SACTCode = xmlDoc.CreateElement("SACTCode");
                                SACTCode.InnerText = getCol0.Current.Value.Trim();
                                list.AppendChild(SACTCode);

                                XPathNodeIterator getCol1 = inDocs.Current.Select("R" + i + "/C1");
                                getCol1.MoveNext();
                                XmlElement SACTDes = xmlDoc.CreateElement("SACTDest");
                                SACTDes.InnerText = getCol1.Current.Value.Trim();
                                list.AppendChild(SACTDes);

                                XPathNodeIterator getCol2 = inDocs.Current.Select("R" + i + "/C2");
                                getCol2.MoveNext();
                                XmlElement ACTCode = xmlDoc.CreateElement("ACTCode");
                                ACTCode.InnerText = getCol2.Current.Value.Trim();
                                list.AppendChild(ACTCode);
                            }
                        }
                        #endregion
                    }
                }
            }
            xmlDoc.Save(sw);
            writer.Write(sw);
            writer.Close();
            fs.Close();
        }

        private static void initFile()
        {
            FileStream fs = new FileStream(pathTemporary, FileMode.Create);
            StreamWriter h = new StreamWriter(fs, System.Text.Encoding.UTF8);
            h.Write("");
            h.Close();
            fs.Close();
        }


        /// <summary>
        /// create database from sql server
        /// </summary>
        /// <returns></returns>
        private static XElement createMinistryFromDB()
        {
            var min = from m in bfAdmin.Ministries select m;
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Ministry")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header in first column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MinistryCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MinistryName"));
            TBody.Add(td);
            foreach (var m in min)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), m.MinistryCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), m.MinistryName));
                TBody.Add(td);
                count++;
            }
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createOrganizationFromDB()
        {
            countCol = 3;
            var org = from o in bfAdmin.Organizations select o;
            XElement herdtable2 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Organization")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Organization Code"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Organization Name"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Ministry Code"));
            TBody.Add(td);
            foreach (var o in org)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), o.OrgCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), o.OrgName),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), o.MinistryCode));
                TBody.Add(td);
                count++;
            }
            herdtable2.Add(countrow);
            herdtable2.Add(TBody);
            return herdtable2;
        }

        private static XElement createProviderFromDB()
        {
            countCol = 3;
            var pro = from p in bfAdmin.Providers select p;
            XElement herdtable3 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Provider")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Provider Code"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Provider Name"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "OrgCode Code"));
            TBody.Add(td);
            foreach (var p in pro)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), p.ProviderCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), p.ProviderName),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), p.OrgCode));
                TBody.Add(td);
                count++;
            }
            herdtable3.Add(countrow);
            herdtable3.Add(TBody);
            return herdtable3;
        }

        private static XElement createActivityCostFromDB()
        {
            countCol = 12;
            var acost = from m in bfAdmin.ActivityCosts select m;
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "ActivityCost")
                     );
            XElement countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 130px;"));
                countrow.Add(countrow2);
            }
            int count = 0;
            XElement TBody = new XElement("TBODY");
            //insert header in first column
            #region
            XElement td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTName"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Unit"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "LabourCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MaterialCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CC_Equipment"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CC_Building"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "IndirectCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ProposedCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CurrentCost"),
                         new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "UnitCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ReferencedCostOrg"));
            #endregion
            TBody.Add(td);
            foreach (var a in acost)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), a.ACTCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), a.ACTName),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), a.Unit),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), a.LabourCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), a.MaterialCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), a.CC_Equipment),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), a.CC_Building),
                     new XElement("TD", new XAttribute("id", "table0_cell_c7_r" + count), a.IndirectCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), a.ProposedCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c9_r" + count), a.CurrentCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c10_r" + count), a.UnitCost),
                     new XElement("TD", new XAttribute("id", "table0_cell_c11_r" + count), a.ReferencedCostOrg));
                TBody.Add(td);
                count++;
            }
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createServiceFromDB()
        {
            countCol = 11;
            var ser = from s in bfAdmin.Services select s;
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Service")
                     );
            XElement countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 120px;"));
                countrow.Add(countrow2);
            }
            int count = 0;
            XElement TBody = new XElement("TBODY");
            //insert header in first column

            XElement td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCName"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCDesc"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "HostCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "StaffRole"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCType"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCObjective"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCSupport"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCoverage"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCStart"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCEnd"));
            TBody.Add(td);
            foreach (var s in ser)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), s.SVCCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), s.SVCName),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), s.SVCDesc),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), s.HostCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), s.StaffRole),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), s.SVCType),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), s.SVCObjective),
                     new XElement("TD", new XAttribute("id", "table0_cell_c7_r" + count), s.SVCSupport),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), s.SVCCoverage),
                     new XElement("TD", new XAttribute("id", "table0_cell_c9_r" + count), s.SVCStart),
                     new XElement("TD", new XAttribute("id", "table0_cell_c10_r" + count), s.SVCEnd));
                TBody.Add(td);
                count++;
            }
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createActivityFromDB()
        {
            countCol = 3;
            var act = from a in bfAdmin.Activities select a;
            XElement herdtable2 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Activity")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTDes"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCode"));
            TBody.Add(td);
            foreach (var a in act)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), a.ACTCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), a.ACTDesc),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), a.SVCCode));
                TBody.Add(td);
                count++;
            }
            herdtable2.Add(countrow);
            herdtable2.Add(TBody);
            return herdtable2;
        }

        private static XElement createSubActivityFromDB()
        {
            countCol = 3;
            var sact = from s in bfAdmin.SubActivities select s;
            XElement herdtable3 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "SubActivity")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SACTCode"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SACTDes"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"));
            TBody.Add(td);
            foreach (var s in sact)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), s.SACTCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), s.SACTDesc),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), s.ACTCode));
                TBody.Add(td);
                count++;
            }
            herdtable3.Add(countrow);
            herdtable3.Add(TBody);
            return herdtable3;
        }

        private static XElement createMaterialFromDB()
        {
            countCol = 7;
            var mat = from m in bfAdmin.Materials select m;
            XElement herdtable4 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Material")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 180px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Material Code"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Material Description"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Unit"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "EstimatedPrice"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "RealPrice"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Service Code"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Note"));
            TBody.Add(td);
            foreach (var m in mat)
            {
                td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), m.MaterialCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), m.MaterialDesc),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), m.Unit),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), m.EstimatedPrice),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), m.RealPrice),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), m.SVCCode),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), m.Note));
                TBody.Add(td);
                count++;
            }
            herdtable4.Add(countrow);
            herdtable4.Add(TBody);
            return herdtable4;
        }

        /// <summary>
        /// create database from html
        /// </summary>
        /// <returns></returns>
        private static XElement createMinistryHtml()
        {
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Ministry")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header in first column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MinistryCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MinistryName"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""));
            TBody.Add(td);
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createOrganizationHtml()
        {
            countCol = 3;
            XElement herdtable2 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Organization")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Organization Code"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Organization Name"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Ministry Code"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""));
            TBody.Add(td);
            herdtable2.Add(countrow);
            herdtable2.Add(TBody);
            return herdtable2;
        }

        private static XElement createProviderHtml()
        {
            countCol = 3;
            XElement herdtable3 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Provider")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Provider Code"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Provider Name"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "OrgCode Code"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""));
            TBody.Add(td);
            herdtable3.Add(countrow);
            herdtable3.Add(TBody);
            return herdtable3;
        }

        private static XElement createActivityCostHtml()
        {
            countCol = 12;
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
            new XAttribute("cellspacing", 0),
            new XAttribute("cellpadding", 0),
            new XAttribute("border", border),
            new XAttribute("id", id),
            new XAttribute("class", classsheet),
            new XAttribute("title", "ActivityCost")
           );
            XElement countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 130px;"));
                countrow.Add(countrow2);
            }
            int count = 0;
            XElement TBody = new XElement("TBODY");
            //insert header in first column
            #region
            XElement td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTName"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Unit"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "LabourCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "MaterialCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CC_Equipment"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CC_Building"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "IndirectCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ProposedCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "CurrentCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "UnitCost"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ReferencedCostOrg"));
            #endregion
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c7_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c9_r" + count), ""));
            TBody.Add(td);
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createServiceHtml()
        {
            countCol = 11;
            XElement herdtable = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Service")
                     );
            XElement countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 120px;"));
                countrow.Add(countrow2);
            }
            int count = 0;
            XElement TBody = new XElement("TBODY");
            //insert header in first column

            XElement td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCName"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCDesc"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "HostCode"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "StaffRole"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCType"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCObjective"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCSupport"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCoverage"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCStart"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCEnd"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c7_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c8_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c9_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c10_r" + count), ""));
            TBody.Add(td);
            herdtable.Add(countrow);
            herdtable.Add(TBody);
            return herdtable;
        }

        private static XElement createActivityHtml()
        {
            countCol = 3;
            var act = from a in bfAdmin.Activities select a;
            XElement herdtable2 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Activity")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTDes"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SVCCode"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""));
            TBody.Add(td);
            herdtable2.Add(countrow);
            herdtable2.Add(TBody);
            return herdtable2;
        }

        private static XElement createSubActivityHtml()
        {
            countCol = 3;
            var sact = from s in bfAdmin.SubActivities select s;
            XElement herdtable3 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "SubActivity")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 200px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SACTCode"),
                           new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "SACTDes"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "ACTCode"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""));
            TBody.Add(td);
            herdtable3.Add(countrow);
            herdtable3.Add(TBody);
            return herdtable3;
        }

        private static XElement createMaterialHtml()
        {
            countCol = 7;
            var mat = from m in bfAdmin.Materials select m;
            XElement herdtable4 = new XElement("TABLE", new XAttribute("style", 100),
                      new XAttribute("cellspacing", 0),
                      new XAttribute("cellpadding", 0),
                      new XAttribute("border", border),
                      new XAttribute("id", id),
                      new XAttribute("class", classsheet),
                      new XAttribute("title", "Material")
                     );
            countrow = new XElement("COLOGROUP");
            for (int i = 0; i < countCol; i++)
            {
                XElement countrow2 = new XElement("COL", new XAttribute("width", 100), new XAttribute("style", "width: 180px;"));
                countrow.Add(countrow2);
            }
            count = 0;
            TBody = new XElement("TBODY");
            //insert header of column
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Material Code"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Material Description"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Unit"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "EstimatedPrice"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "RealPrice"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Service Code"),
                          new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + ++count),
                              new XAttribute("class", "styleBold styleCenter"),
                              new XAttribute("style", "background-color: rgb(192, 192, 192)"),
                              "Note"));
            TBody.Add(td);
            td = new XElement("TR", new XAttribute("style", "height: 25px;"),
                     new XElement("TD", new XAttribute("id", "table0_cell_c0_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c1_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c2_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c3_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c4_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c5_r" + count), ""),
                     new XElement("TD", new XAttribute("id", "table0_cell_c6_r" + count), ""));
            TBody.Add(td);
            herdtable4.Add(countrow);
            herdtable4.Add(TBody);
            return herdtable4;
        }

    }
}