using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Xml;

namespace Spreadsheet
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IXMLService" in both code and config file together.
    [ServiceContract]
    public interface IXMLService
    {
        [OperationContract(Name = "GetXMLSheetCode")]
        [WebGet(UriTemplate = "/xmlsheet/getcode?pk={pk}&fk={fk}&select={select}")]
        XmlElement getXMLSheet(string select, string pk, string fk);

        [OperationContract(Name = "GetXMLSheet")]
        [WebGet(UriTemplate = "/xmlsheet/getonesheet?group={group}&sheet={sheet}")]
        XmlElement getXMLSheet(string group, string sheet);
    }
}
