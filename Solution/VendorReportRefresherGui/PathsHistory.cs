using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Configuration;

namespace GDPptGeneratorUI
{
    internal class PathsHistory
    {
        private int _maxHistoryLength;

        private List<string> _budgetPaths;
        private List<string> _forecastPaths;
        private List<string> _superDettagliPaths;
        private List<string> _ranRatePaths;
        private List<string> _destFolderPaths;


        private string _xmlFilePath;

        public List<string> BudgetPaths { get => _budgetPaths; set => _budgetPaths = value; }
        public List<string> ForecastPaths { get => _forecastPaths; set => _forecastPaths = value; }
        public List<string> SuperDettagliPaths { get => _superDettagliPaths; set => _superDettagliPaths = value; }
        public List<string> RanRatePaths { get => _ranRatePaths; set => _ranRatePaths = value; }
        public List<string> DestFolderPaths { get => _destFolderPaths; set => _destFolderPaths = value; }


        private const string XML_KEY_BUDGET = "BudgetFilePaths";


        public PathsHistory(string xmlFilePath)
        {
            string maxHistoryLengthStr = ConfigurationManager.AppSettings.Get("HistoryFilePathMaxLength");

            if (string.IsNullOrEmpty(maxHistoryLengthStr) || !int.TryParse(maxHistoryLengthStr, out _maxHistoryLength))
            {
                _maxHistoryLength = 50;
            }

            XmlDocument doc = LoadXmlDocument(_xmlFilePath);

            _budgetPaths = fillListFromXmlForHystoryType(doc, XML_KEY_BUDGET);
            _forecastPaths = fillListFromXmlForHystoryType(doc, "Forecast");
            _superDettagliPaths = fillListFromXmlForHystoryType(doc, "SuperDettagli");
            _ranRatePaths = fillListFromXmlForHystoryType(doc, "RanRate");
            _destFolderPaths = fillListFromXmlForHystoryType(doc, "DestinationFolder");
            return;

            var budgetHistoryNodeList = doc.SelectNodes("/FileHistory/Budget/FilePath");
            _budgetPaths = new List<string>();
            foreach (XmlNode node in budgetHistoryNodeList)
            { _budgetPaths.Add(node.InnerText); }            

            var forecastHistoryNodeList = doc.SelectNodes("/FileHistory/Forecast/FilePath");
            _forecastPaths = new List<string>();
            foreach (XmlNode node in forecastHistoryNodeList)
            { _forecastPaths.Add(node.InnerText); }

            var superDettagliHistoryNodeList = doc.SelectNodes("/FileHistory/SuperDettagli/FilePath");
            _superDettagliPaths = new List<string>();
            foreach (XmlNode node in superDettagliHistoryNodeList)
            { _superDettagliPaths.Add(node.InnerText); }

            var ranRateHistoryNodeList = doc.SelectNodes("/FileHistory/RanRate/FilePath");
            _ranRatePaths = new List<string>();
            foreach (XmlNode node in ranRateHistoryNodeList)
            { _ranRatePaths.Add(node.InnerText); }

            var destFolderHistoryNodeList = doc.SelectNodes("/FileHistory/DestinationFolder/Path");
            _destFolderPaths = new List<string>();
            foreach (XmlNode node in destFolderHistoryNodeList)
            { _destFolderPaths.Add(node.InnerText); }
        }

        private List<string> fillListFromXmlForHystoryType(XmlDocument doc, string hystoryElementType)
        {
            var xmlPath = $"/FileHistory/{hystoryElementType}/FilePath";
            var historyNodeList = doc.SelectNodes(xmlPath);


            var hystoryElements = new List<string>();
            foreach (XmlNode node in historyNodeList)
            { hystoryElements.Add(node.InnerText); }
            return hystoryElements;
        }

        private XmlDocument LoadXmlDocument(string pXmlFilePath)
        {
            XmlDocument doc = new XmlDocument();

            if (File.Exists(_xmlFilePath))
            {
                doc.Load(_xmlFilePath);
            }

            return doc;
        }

        public void AddPathsHistory(string budgetPath, string forecastPath, string superDettagliPath, string ranRatePath, string destinationFolderPath)
        {
            XmlDocument doc = new XmlDocument();

            XmlElement fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            fileHistoryElement.AppendChild(updateHystoryItems(doc, budgetPath, _budgetPaths, XML_KEY_BUDGET));

            //XmlElement budgetPathElement = doc.CreateElement(XML_KEY_BUDGET);
            //fileHistoryElement.AppendChild(budgetPathElement);

            XmlElement forecastPathElement = doc.CreateElement("ForecastPaths");
            fileHistoryElement.AppendChild(forecastPathElement);

            XmlElement superDettagliPathElement = doc.CreateElement("SuperDettagliPaths");
            fileHistoryElement.AppendChild(superDettagliPathElement);

            XmlElement ranRatePathElement = doc.CreateElement("RanRatePaths");
            fileHistoryElement.AppendChild(ranRatePathElement);

            XmlElement destFoldersElement = doc.CreateElement("DestFolders");
            fileHistoryElement.AppendChild(destFoldersElement);



            //if (_maxHistoryLength > 0)
            //{
            //    //Controllers
            //    XmlElement newControllerElement = doc.CreateElement("FilePath");
            //    newControllerElement.InnerText = controllerPath;
            //    controllersElement.AppendChild(newControllerElement);
            //}

            //int itemsCount = 1;

            //foreach (string exsistingControllerPath in _controllerPaths)
            //{
            //    if (!exsistingControllerPath.Equals(controllerPath, StringComparison.OrdinalIgnoreCase))
            //    {
            //        if (itemsCount >= _maxHistoryLength)
            //            break;

            //        XmlElement exsistingControllerElement = doc.CreateElement("FilePath");
            //        exsistingControllerElement.InnerText = exsistingControllerPath;
            //        controllersElement.AppendChild(exsistingControllerElement);

            //        itemsCount++;
            //    }
            //}





            //itemsCount = 1;

            //if (_maxHistoryLength > 0)
            //{
            //    //Reports
            //    XmlElement newDestFolderElement = doc.CreateElement("Path");
            //    newDestFolderElement.InnerText = destinationFolderPath;
            //    destFoldersElement.AppendChild(newDestFolderElement);
            //}

            //foreach (string exsistingDestFolders in _destFolderPaths)
            //{
            //    if (!exsistingDestFolders.Equals(destinationFolderPath, StringComparison.OrdinalIgnoreCase))
            //    {
            //        if (itemsCount >= _maxHistoryLength)
            //            break;

            //        XmlElement exsistingDestFolderElement = doc.CreateElement("Path");
            //        exsistingDestFolderElement.InnerText = exsistingDestFolders;
            //        destFoldersElement.AppendChild(exsistingDestFolderElement);

            //        itemsCount++;
            //    }
            //}


            doc.Save(_xmlFilePath);
        }


        private XmlElement updateHystoryItems(XmlDocument doc, string firstPostion, List<string> currentItems, string key)
        {
            XmlElement currentPathElement = doc.CreateElement(key);

            if (_maxHistoryLength > 0)
            {
                //Controllers
                XmlElement newItem = doc.CreateElement("FilePath");
                newItem.InnerText = firstPostion;
                currentPathElement.AppendChild(newItem);
            }

            int itemsCount = 1;

            foreach (string exsistingItem in currentItems)
            {
                if (!exsistingItem.Equals(firstPostion, StringComparison.OrdinalIgnoreCase))
                {
                    if (itemsCount >= _maxHistoryLength)
                        break;

                    XmlElement exsistingItmeElementToInsertAgain = doc.CreateElement("FilePath");
                    exsistingItmeElementToInsertAgain.InnerText = exsistingItem;
                    currentPathElement.AppendChild(exsistingItmeElementToInsertAgain);

                    itemsCount++;
                }
            }

            return currentPathElement;
        }

        public void ClearHistory()
        {
            XmlDocument doc = new XmlDocument();

            XmlElement fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            XmlElement budgetPathElement = doc.CreateElement(XML_KEY_BUDGET);
            fileHistoryElement.AppendChild(budgetPathElement);

            XmlElement forecastPathElement = doc.CreateElement("ForecastPaths");
            fileHistoryElement.AppendChild(forecastPathElement);

            XmlElement superDettagliPathElement = doc.CreateElement("SuperDettagliPaths");
            fileHistoryElement.AppendChild(superDettagliPathElement);

            XmlElement ranRatePathElement = doc.CreateElement("RanRatePaths");
            fileHistoryElement.AppendChild(ranRatePathElement);

            XmlElement destFoldersElement = doc.CreateElement("DestFolders");
            fileHistoryElement.AppendChild(destFoldersElement);

            //XmlDocument doc = new XmlDocument();

            //XmlElement fileHistoryElement = doc.CreateElement("FileHistory");
            //doc.AppendChild(fileHistoryElement);

            //XmlElement controllersElement = doc.CreateElement("Controllers");
            //fileHistoryElement.AppendChild(controllersElement);

            //XmlElement reportsElement = doc.CreateElement("Reports");
            //fileHistoryElement.AppendChild(reportsElement);

            //XmlElement destFoldersElement = doc.CreateElement("DestFolders");
            //fileHistoryElement.AppendChild(reportsElement);

            doc.Save(_xmlFilePath);
        }
    }
}
