using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System.Linq;

namespace FilesEditor.Steps
{
    internal class Step_CreaLista_ItemsToExportAsImage : StepBase
    {
        public Step_CreaLista_ItemsToExportAsImage(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_CreaLista_ItemsToExportAsImage", Context);
            creaListaImmaginiDaEsportare();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void creaListaImmaginiDaEsportare()
        {
            var ePPlusHelper = GetHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);

            var imageIds = Context.SildeToGenerate.SelectMany(_ => _.Contents).Distinct().ToList();
            foreach (var imageId in imageIds)
            {
                ThrowExpetionsForMissingWorksheet(ePPlusHelper, imageId, FileTypes.DataSource);

                var printArea = ePPlusHelper.GetString(imageId, Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW, Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL);
                // check sul campo "Print Area"
                if (string.IsNullOrWhiteSpace(printArea))
                {
                    throw new ManagedException(
                        filePath: ePPlusHelper.FilePathInUse,
                        fileType: FileTypes.DataSource,
                        //
                        worksheetName: imageId,
                        cellRow: Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW,
                        cellColumn: Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL,
                        valueHeader: ValueHeaders.None,
                        value: null,
                        //
                        errorType: ErrorTypes.MissingValue,
                        userMessage: $"Print area not defined for sheet '{imageId}'."
                        );
                }

                Context.ItemsToExportAsImage.Add(new ItemToExport { ImageId = imageId, Sheet = imageId, PrintArea = printArea, });
            }
        }
    }
}