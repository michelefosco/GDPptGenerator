using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Xml;

namespace PptGeneratorGUI
{
    internal class PathsHistory
    {
        private int _maxHistoryLength;
        private string _xmlFilePath;

        //BudgetPaths
        private List<string> _budgetPaths;
        public List<string> BudgetPaths { get => _budgetPaths; set => _budgetPaths = value; }
        private const string XML_KEY_BUDGET = "BudgetFilePaths";


        //ForecastPaths
        private List<string> _forecastPaths;
        public List<string> ForecastPaths { get => _forecastPaths; set => _forecastPaths = value; }
        private const string XML_KEY_FORECAST = "Forecast";


        //SuperDettagliPaths
        private List<string> _superDettagliPaths;
        public List<string> SuperDettagliPaths { get => _superDettagliPaths; set => _superDettagliPaths = value; }
        private const string XML_KEY_SUPERDETTAGLI = "SuperDettagli";


        //RunRatePaths
        private List<string> _runRatePaths;
        public List<string> RunRatePaths { get => _runRatePaths; set => _runRatePaths = value; }
        private const string XML_KEY_RUNRATE = "RunRate";


        //DestFolderPaths
        private List<string> _destFolderPaths;
        public List<string> DestFolderPaths { get => _destFolderPaths; set => _destFolderPaths = value; }
        private const string XML_KEY_DESTINATIONFOLDER = "DestinationFolder";


        public PathsHistory(string xmlFilePath)
        {
            _xmlFilePath = xmlFilePath;

            var maxHistoryLengthStr = ConfigurationManager.AppSettings.Get("HistoryFilePathMaxLength");

            if (string.IsNullOrEmpty(maxHistoryLengthStr) || !int.TryParse(maxHistoryLengthStr, out _maxHistoryLength))
            {
                _maxHistoryLength = 50; // Default value if not set or invalid
            }

            XmlDocument doc = LoadXmlDocument(_xmlFilePath);

            _budgetPaths = fillListFromXmlForHystoryType(doc, XML_KEY_BUDGET);
            _forecastPaths = fillListFromXmlForHystoryType(doc, XML_KEY_FORECAST);
            _superDettagliPaths = fillListFromXmlForHystoryType(doc, XML_KEY_SUPERDETTAGLI);
            _runRatePaths = fillListFromXmlForHystoryType(doc, XML_KEY_RUNRATE);
            _destFolderPaths = fillListFromXmlForHystoryType(doc, XML_KEY_DESTINATIONFOLDER);
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
            var doc = new XmlDocument();

            if (File.Exists(_xmlFilePath))
            {
                doc.Load(_xmlFilePath);
            }

            return doc;
        }

        public void AddPathsHistory(string budgetPath, string forecastPath, string superDettagliPath, string runRatePath, string destinationFolderPath)
        {
            var doc = new XmlDocument();

            var fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            fileHistoryElement.AppendChild(updateHystoryItems(doc, budgetPath, _budgetPaths, XML_KEY_BUDGET));
            fileHistoryElement.AppendChild(updateHystoryItems(doc, forecastPath, _forecastPaths, XML_KEY_FORECAST));
            fileHistoryElement.AppendChild(updateHystoryItems(doc, superDettagliPath, _superDettagliPaths, XML_KEY_SUPERDETTAGLI));
            fileHistoryElement.AppendChild(updateHystoryItems(doc, runRatePath, _runRatePaths, XML_KEY_RUNRATE));
            fileHistoryElement.AppendChild(updateHystoryItems(doc, destinationFolderPath, _destFolderPaths, XML_KEY_DESTINATIONFOLDER));

            doc.Save(_xmlFilePath);
        }


        private XmlElement updateHystoryItems(XmlDocument doc, string firstPostion, List<string> currentItems, string key)
        {
            var currentPathElement = doc.CreateElement(key);

            if (_maxHistoryLength > 0)
            {
                var newItem = doc.CreateElement("FilePath");
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

                    var exsistingItmeElementToInsertAgain = doc.CreateElement("FilePath");
                    exsistingItmeElementToInsertAgain.InnerText = exsistingItem;
                    currentPathElement.AppendChild(exsistingItmeElementToInsertAgain);

                    itemsCount++;
                }
            }

            return currentPathElement;
        }

        public void ClearHistory()
        {
            var doc = new XmlDocument();

            var fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            fileHistoryElement.AppendChild(doc.CreateElement(XML_KEY_BUDGET));
            fileHistoryElement.AppendChild(doc.CreateElement(XML_KEY_FORECAST));
            fileHistoryElement.AppendChild(doc.CreateElement(XML_KEY_SUPERDETTAGLI));
            fileHistoryElement.AppendChild(doc.CreateElement(XML_KEY_RUNRATE));
            fileHistoryElement.AppendChild(doc.CreateElement(XML_KEY_DESTINATIONFOLDER));

            doc.Save(_xmlFilePath);
        }
    }
}
