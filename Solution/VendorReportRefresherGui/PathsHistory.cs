using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Configuration;

namespace VendorReportRefresher
{
    internal class PathsHistory
    {
        private List<string> _controllerPaths;
        private List<string> _reportPaths;
        private List<string> _destFolderPaths;
        private int _maxHistoryLength;

        public List<string> ControllerPaths { get => _controllerPaths; set => _controllerPaths = value; }
        public List<string> ReportPaths { get => _reportPaths; set => _reportPaths = value; }
        public List<string> DestFolderPaths { get => _destFolderPaths; set => _destFolderPaths = value; }

        private string _xmlFilePath;

        public PathsHistory(string xmlFilePath)
        {
            string maxHistoryLengthStr = ConfigurationManager.AppSettings.Get("HistoryFilePathMaxLength");

            if (string.IsNullOrEmpty(maxHistoryLengthStr) || !int.TryParse(maxHistoryLengthStr, out _maxHistoryLength))
            {
                _maxHistoryLength = 100;
            }

            _xmlFilePath = xmlFilePath;
            _controllerPaths = new List<string>();
            _reportPaths = new List<string>();

            XmlDocument doc = LoadXmlDocument(_xmlFilePath);

            var controllerHistoryNodeList = doc.SelectNodes("/FileHistory/Controllers/FilePath");
            var reportHistoryNodeList = doc.SelectNodes("/FileHistory/Reports/FilePath");
            var destFolderHistoryNodeList = doc.SelectNodes("/FileHistory/DestFolders/Path");

            _controllerPaths = new List<string>();
            _reportPaths = new List<string>();
            _destFolderPaths = new List<string>();

            foreach (XmlNode controllerNode in controllerHistoryNodeList)
            {
                _controllerPaths.Add(controllerNode.InnerText);
            }

            foreach(XmlNode reportNode in reportHistoryNodeList)
            {
                _reportPaths.Add(reportNode.InnerText);
            }

            foreach (XmlNode folderNode in destFolderHistoryNodeList)
            {
                _destFolderPaths.Add(folderNode.InnerText);
            }
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

        public void AddPathsHistory(string controllerPath, string reportPath, string destFolderPath)
        {
            XmlDocument doc = new XmlDocument();

            XmlElement fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            XmlElement controllersElement = doc.CreateElement("Controllers");
            fileHistoryElement.AppendChild(controllersElement);

            XmlElement reportsElement = doc.CreateElement("Reports");
            fileHistoryElement.AppendChild(reportsElement);

            XmlElement destFoldersElement = doc.CreateElement("DestFolders");
            fileHistoryElement.AppendChild(destFoldersElement);

            if (_maxHistoryLength > 0)
            {
                //Controllers
                XmlElement newControllerElement = doc.CreateElement("FilePath");
                newControllerElement.InnerText = controllerPath;
                controllersElement.AppendChild(newControllerElement);
            }

            int itemsCount = 1;

            foreach(string exsistingControllerPath in _controllerPaths)
            {
                if (!exsistingControllerPath.Equals(controllerPath,StringComparison.OrdinalIgnoreCase))
                {
                    if (itemsCount >= _maxHistoryLength)
                        break;

                    XmlElement exsistingControllerElement = doc.CreateElement("FilePath");
                    exsistingControllerElement.InnerText = exsistingControllerPath;
                    controllersElement.AppendChild(exsistingControllerElement);

                    itemsCount++;
                }
            }

            itemsCount = 1;

            if (_maxHistoryLength > 0)
            {
                //Reports
                XmlElement newReportElement = doc.CreateElement("FilePath");
                newReportElement.InnerText = reportPath;
                reportsElement.AppendChild(newReportElement);
            }

            foreach (string exsistingReportPath in _reportPaths)
            {
                if (!exsistingReportPath.Equals(reportPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (itemsCount >= _maxHistoryLength)
                        break;

                    XmlElement exsistingReportElement = doc.CreateElement("FilePath");
                    exsistingReportElement.InnerText = exsistingReportPath;
                    reportsElement.AppendChild(exsistingReportElement);
                 
                    itemsCount++;
                }
            }

            itemsCount = 1;

            if (_maxHistoryLength > 0)
            {
                //Reports
                XmlElement newDestFolderElement = doc.CreateElement("Path");
                newDestFolderElement.InnerText = destFolderPath;
                destFoldersElement.AppendChild(newDestFolderElement);
            }

            foreach (string exsistingDestFolders in _destFolderPaths)
            {
                if (!exsistingDestFolders.Equals(destFolderPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (itemsCount >= _maxHistoryLength)
                        break;

                    XmlElement exsistingDestFolderElement = doc.CreateElement("Path");
                    exsistingDestFolderElement.InnerText = exsistingDestFolders;
                    destFoldersElement.AppendChild(exsistingDestFolderElement);

                    itemsCount++;
                }
            }

            doc.Save(_xmlFilePath);
        }

        public void ClearHistory()
        {
            XmlDocument doc = new XmlDocument();

            XmlElement fileHistoryElement = doc.CreateElement("FileHistory");
            doc.AppendChild(fileHistoryElement);

            XmlElement controllersElement = doc.CreateElement("Controllers");
            fileHistoryElement.AppendChild(controllersElement);

            XmlElement reportsElement = doc.CreateElement("Reports");
            fileHistoryElement.AppendChild(reportsElement);

            XmlElement destFoldersElement = doc.CreateElement("DestFolders");
            fileHistoryElement.AppendChild(reportsElement);

            doc.Save(_xmlFilePath);
        }
    }
}
