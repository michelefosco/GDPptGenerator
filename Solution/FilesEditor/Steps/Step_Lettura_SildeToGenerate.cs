using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_Lettura_SildeToGenerate : StepBase
    {
        public Step_Lettura_SildeToGenerate(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            getDataFromDataSourceFile();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
        
        private  void getDataFromDataSourceFile()
        {
            var dataSourceTemplateFile = Path.Combine(Context.SourceFilesFolder, FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            var slidesToGenerate = getSildeToGenerate(ePPlusHelper, Context.Configurazione);
            Context.SildeToGenerate = slidesToGenerate;            
        }
        private  List<SlideToGenerate> getSildeToGenerate(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            var printableWorksheets = ePPlusHelper.GetWorksheetNames().Where(n => n.StartsWith(WorksheetNames.DATA_SOURCE_TEMPLATE_PRINTABLE_WORKSHEET_NAME_PREFIX, StringComparison.InvariantCultureIgnoreCase)).ToList();

            //todo: lettura dei workshitt names con elementi stampabili

            var slideToGenerateList = new List<SlideToGenerate>();

            var rigaCorrente = configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_FIRST_DATA_ROW;
            while (true)
            {
                var outputFileName = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL);
                var title = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL);
                var content1 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL);
                var content2 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_2_COL);
                var content3 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_3_COL);
                var layout = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL);

                // mi fermo quando la riga è competamente a null
                if (allNulls(outputFileName, title, content1, content2, content3, layout))
                { break; }

                //verifico i campi obbligatori
                // check sul campo "Powerpoint File"
                ManagedException.ThrowIfMissingMandatoryValue(outputFileName, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL,
                    ValueHeaders.TableName);
                // check sul campo "Title"
                ManagedException.ThrowIfMissingMandatoryValue(title, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL,
                    ValueHeaders.SlideTitle);

                // check sul campo "Content 1"
                ManagedException.ThrowIfMissingMandatoryValue(content1, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL,
                    ValueHeaders.SlideTitle);

                var contents = new List<string>() { content1 };
                if (!string.IsNullOrWhiteSpace(content2)) { contents.Add(content2); }
                if (!string.IsNullOrWhiteSpace(content3)) { contents.Add(content3); }
                // Verifico che i valori usati in contents siano validi (ovvero esistano foglio con quel nome)
                foreach (var item in contents)
                {
                    if (!printableWorksheets.Any(n => n.Equals(item, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        //todo: Sollevare eccezione Managed
                        throw new Exception("Elemento da stampare non valido");
                    }
                }

                // con più di un contenuto, il layout diventa obbligatorio
                if (contents.Count > 1)
                {
                    // check sul campo "Layout"
                    ManagedException.ThrowIfMissingMandatoryValue(layout, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                        configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL,
                        ValueHeaders.SlideLayout);
                }
                //todo: chiedere info a Francesco su questo uso del default
                if (layout == null)
                { layout = LayoutTypes.Horizontal.ToString(); }


                if (Enum.TryParse(layout, out LayoutTypes layoutType))
                {
                    // ok
                }
                else
                {
                    // Sollevare eccezione Managed
                    //todo:
                    throw new Exception("Tipo layout sconosciuto");
                }


                // aggiungo la slide alla lista di quelle lette
                slideToGenerateList.Add(new SlideToGenerate
                {
                    OutputFileName = outputFileName,
                    Title = title,
                    LayoutType = layoutType,
                    Contents = contents
                });

                // passo alla riga successiva
                rigaCorrente++;
            }

            return slideToGenerateList;
        }
    }
}