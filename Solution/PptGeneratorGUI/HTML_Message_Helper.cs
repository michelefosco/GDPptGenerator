using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PptGeneratorGUI
{
    public class HTML_Message_Helper
    {
        public const string _newlineHTML = @"<BR />";
        public const string _boldHTML = @"<B>{0}</B>";
        public const string _hyperlinkHTML = @"<a style=""color: blue;"" href=""{0}"">{1}</a>";
        public const string _tabHTML = "&nbsp;&nbsp;&nbsp;";
        public const string _spaceHTML = "&nbsp;";
        public const string _redTextHTML = @"<span class=""red"">{0}</span>";
        public const string _greenTextHTML = @"<span class=""green"">{0}</span>";
        public const string _tableHTML = "<table>\r\n{0}\r\n</table>";
        public const string _trHTML = "  <tr>\r\n{0}\r\n  </tr>";
        public const string _tdHTML = "    <td>{0}</td>";
        public const string _invisibleSpanHTML = "<span id=\"invisibleSpan\" style=\"display: none;\">{0}</span>";
        public const string _moreDetailLink = @"<a href=""#"" style=""color: blue;"" onclick=""document.getElementById('invisibleSpan').style.display = 'inline'"">{0}</a>";
        public const string _deleteFileHyperlinkHTML = @"<a style=""color: red;"" href=""{0}"">{1}</a>";
        public const string _htmlBody = @"<html>
	                                        <head>
	                                        <style type=""text/css"">
		                                        body {{ font: 11px sans-serif; }}
		                                        table {{ font: 11px sans-serif; }}
		                                        th, td {{ padding-right: 10px; }}
		                                        .red {{ font-family: sans-serif; color: red; font-size:13px; }}
		                                        .green {{ font-family: sans-serif; color: green; font-size:13px; }}
	                                        </style> 
	                                        </head>
	                                        <body>{0}</body>
                                        </html>";


        public static string StringToHTML(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }

            return str.Replace("\t", _tabHTML)
                .Replace(" ", _spaceHTML)
                .Replace("\r\n", _newlineHTML)
                .Replace("\n", _newlineHTML);
        }
        public static string GetHTMLHyperLink(string url, string value)
        {
            return string.Format(_hyperlinkHTML, GetURLMarker(url), value);
        }
        public static string GetHTMLHyperLinkSetAsImput(string url, string value)
        {
            return string.Format(_hyperlinkHTML, GetURLMarkerSetAsImput(url), value);
        }
        public static string GetHTMLDeleteFileHyperLink(string url)
        {
            return string.Format(_deleteFileHyperlinkHTML, GetURLMarkerDelete(url), "(Clicca quì per cancellare il file)");
        }
        public static string GetHTMLMoreDetailLink(string caption)
        {
            return string.Format(_moreDetailLink, caption);
        }
        public static string GetHTMLBold(string str)
        {
            return string.Format(_boldHTML, str);
        }
        public static string GetHTMLRedText(string str)
        {
            return string.Format(_redTextHTML, str);
        }
        public static string GetHTMLGreenText(string str)
        {
            return string.Format(_greenTextHTML, str);
        }
        public static string GetHTMLTable(string innerTableHTML)
        {
            return string.Format(_tableHTML, innerTableHTML);
        }
        public static string GetHTMLTableRow(string innerRowHTML)
        {
            return string.Format(_trHTML, innerRowHTML);
        }
        public static string GetHTMLTableCell(string innerCellHTML)
        {
            return string.Format(_tdHTML, innerCellHTML);
        }
        public static string GetHTMLTableRowWithCells(string cell1Value, string cell2Value)
        {
            return GetHTMLTableRow(GetHTMLTableCell(cell1Value) + GetHTMLTableCell(GetHTMLBold(cell2Value)));
        }
        public static string GetURLMarker(string url)
        {
            return "file-" + url;
        }
        public static string GetURLMarkerSetAsImput(string url)
        {
            return "input-" + url;
        }
        public static string GetURLMarkerDelete(string url)
        {
            return "filedelete-" + url;
        }
        public static string GetInvisibleSPAN(string innerHtml)
        {
            return string.Format(_invisibleSpanHTML, innerHtml);
        }
        public static string GetHTMLForExpetion(Exception ex)
        {
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Error:"));
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += StringToHTML(ex.Message);
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;

            htmlErrorMessage += GetInvisibleErrorDetails(ex);

            return htmlErrorMessage;
        }
        public static string GetHTMLForExpetion(ManagedException mEx)
        {
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Error:"));
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += StringToHTML(GetHTMLBold(mEx.UserMessage));

            if (!string.IsNullOrEmpty(mEx.FilePath))
            {
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(mEx.FilePath, mEx.FilePath);
                //htmlErrorMessage += _spaceHTML;
                //htmlErrorMessage += GetHTMLDeleteFileHyperLink(mEx.PercorsoFile);
            }
            switch (mEx.FileType)
            {
                case FileTypes.Undefined:
                case FileTypes.Directory:
                    break;

                default:
                    htmlErrorMessage += _newlineHTML;
                    htmlErrorMessage += _newlineHTML;
                    htmlErrorMessage += StringToHTML("File types: ") + GetHTMLBold(mEx.FileType.ToString());
                    break;

                    //case TipologiaCartelle.FileDiTipo2:
                    //    htmlErrorMessage += _newlineHTML;
                    //    htmlErrorMessage += _newlineHTML;
                    //    htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(SelectedControllerFilePath, SelectedControllerFilePath);
                    //    break;
            }

            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;

            //Tabella con dati aggiuntivi dell'errore
            string tableHTML = GetHTMLTableRowWithCells("Error type: ", mEx.ErrorType.GetEnumDescription());
            tableHTML += GetHTMLTableRowWithCells("File type:", mEx.FileType.GetEnumDescription());

            if (!string.IsNullOrEmpty(mEx.WorksheetName))
            {
                tableHTML += GetHTMLTableRowWithCells("Worksheet name:", mEx.WorksheetName);
            }

            if (mEx.CellColumn.HasValue && mEx.CellRow.HasValue)
            {
                tableHTML += GetHTMLTableRowWithCells("Cell:", $"{((ColumnIDS)mEx.CellColumn).ToString()}{mEx.CellRow.ToString()}");
            }
            else
            {
                if (mEx.CellColumn.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Column:", ((ColumnIDS)mEx.CellColumn).ToString());
                }

                if (mEx.CellRow.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Row:", mEx.CellRow.ToString());
                }
            }

            //if (mEx.NomeDatoErrore != NomiDatoErrore.None)
            //{
            //    tableHTML += GetHTMLTableRowWithCells("Errore sul dato:", mEx.NomeDatoErrore.GetEnumDescription());
            //}

            if (!string.IsNullOrEmpty(mEx.Value))
            {
                tableHTML += GetHTMLTableRowWithCells("Value:", mEx.Value);
            }

            htmlErrorMessage += GetHTMLTable(tableHTML);

            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += GetInvisibleErrorDetails(mEx);

            return htmlErrorMessage;
        }
        public static string GetHTMLForBody(string message)
        {
            string htmlMessage = string.Format(_htmlBody, message);
            return htmlMessage;
        }
        public static string GetInvisibleErrorDetails(Exception ex)
        {
            return GetInvisibleSPAN(StringToHTML($"Full error:\r\n{ex}"));
        }
        public static string GeneraHtmlPerWarning(List<string> warnings)
        {
            if (warnings == null || !warnings.Any())
            { return string.Empty; }

            string outputMessage = "";
            outputMessage += _newlineHTML + _newlineHTML + _newlineHTML;
            outputMessage += GetHTMLBold($"Warnings were reported during the processing:");

            //Elenco dei warnings
            outputMessage += "<UL>";
            foreach (var warning in warnings)
            {
                outputMessage += $"<li>{warning}</li>";
            }
            outputMessage += "</UL>";

            return outputMessage;
        }
    }
}
