using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;


namespace ProviderDashboards
{
    public class XMLSettings
    {
        String metricsFromXml = "";
        String dashboardFromXml = "";
        List<String> providersFromXml = new List<String>();

        public void WriteConfigFile(List<String> dataLocations, List<String> providers)
        {
            XDocument doc =
                new XDocument(
                    new XElement("Settings",
                        new XElement("DataLocations",
                            new XElement("MetricsFolder", dataLocations[0]),
                            new XElement("DashboardFile", dataLocations[1])),
                        new XElement("Providers",
                            providers.Select(x => new XElement("Name", x)))));
            doc.Save("Settings.xml");
        }
       
        public void ReadConfigFile(String settingsFile, ref List<String> dataLocations, ref List<String> providers)
        {
            //clear providers variable from form1 so that we dont havea huge list with duplicates
            providers.Clear();
            XmlTextReader reader = new XmlTextReader(settingsFile);
            String previousNodeName = "";

            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // this node is an element
                        break;
                    case XmlNodeType.Text: //display the text in each element
                       if (previousNodeName == "MetricsFolder")
                            {
                                dataLocations.Add(reader.Value);
                            }
                            if (previousNodeName == "DashboardFile")
                            {
                                dataLocations.Add(reader.Value);
                            }
                            if (previousNodeName == "Name")
                            {
                                providers.Add(reader.Value);
                            }
                        break;
                    case XmlNodeType.EndElement: //display the end of the element
                        break;
                }
                previousNodeName = reader.Name;

            }
            reader.Close();

        }

        
    }
}
