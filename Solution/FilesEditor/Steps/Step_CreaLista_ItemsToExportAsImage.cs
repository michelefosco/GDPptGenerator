using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Linq;

namespace FilesEditor.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_CreaLista_ItemsToExportAsImage : StepBase
    {
        internal override string StepName => "Step_CreaLista_ItemsToExportAsImage";

        public Step_CreaLista_ItemsToExportAsImage(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            CreaListaImmaginiDaEsportare();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void CreaListaImmaginiDaEsportare()
        {
            var imageIds = Context.SildeToGenerate.SelectMany(_ => _.Contents).Distinct().ToList();
            foreach (var imageId in imageIds)
            {
                EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(Context.DataSourceEPPlusHelper, imageId, FileTypes.DataSource);

                // ImageId coincide con WorksheetName
                var printArea = Context.DataSourceEPPlusHelper.GetString(imageId, Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW, Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL);
                // check sul campo "Print Area"
                if (string.IsNullOrWhiteSpace(printArea))
                {
                    throw new ManagedException(
                        filePath: Context.DataSourceEPPlusHelper.FilePathInUse,
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
                if(!Context.DataSourceEPPlusHelper.IsValidAddress(printArea))
                {
                    throw new ManagedException(
                        filePath: Context.DataSourceEPPlusHelper.FilePathInUse,
                        fileType: FileTypes.DataSource,
                        //
                        worksheetName: imageId,
                        cellRow: Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW,
                        cellColumn: Context.Configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL,
                        valueHeader: ValueHeaders.None,
                        value: printArea,
                        //
                        errorType: ErrorTypes.InvalidValue,
                        userMessage: $"Print area '{printArea}' is not a valid address"
                        );
                }

                //todo: confermare comportamento i grafici...
                printArea = Context.DataSourceEPPlusHelper.ReduceAreaToDimensionEnd(imageId, printArea);

                Context.ItemsToExportAsImage.Add(
                        new ItemToExport(
                            workSheetName: imageId,
                            printArea: printArea,
                            imageId: imageId,
                            imageFilePath: GetTmpFolderImagePathByImageId(Context.TmpFolder, imageId)
                            )
                    );
            }
        }
    }
}