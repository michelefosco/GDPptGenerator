using FilesEditor.Constants;
using FilesEditor.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesEditor
{
    public class Info
    {
        public static ValidaSourceFilesOutput ValidaSourceFiles(ValidaSourceFilesInput validaSourceFilesInput)
        {
            // Lettura Opzione dal datasource

            var dataSourceTemplateFile = Path.Combine(validaSourceFilesInput.TemplatesFolder, Constants.FileNames.DATA_SOURCE_FILENAME);
            var opzioniUtente = GetOpzioniUtente(dataSourceTemplateFile);

            // lettura info da 1° file

            // lettura info da 2° file

            // lettura info da 3° file

            // lettura info da 4° file

            // verifica applicabilità filtri

            // completa Valori sele

            return new ValidaSourceFilesOutput
            {
                OpzioniUtente = opzioniUtente
            };
        }


        private static OpzioniUtente GetOpzioniUtente(string filePath)
        {
            //todo leggere da file
            var filtriPossibili = new List<FilterItems>();
            filtriPossibili.Add(new FilterItems
            {
                Tabella = "SUPERDETTAGLI",
                Campo = "ProjType Cluster 2_",
                ValoriPossibili = new List<string>(),
                ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED }
            });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BussinessArea Cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BusinessArea_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "ProjType_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Proj type cluster 2", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Business", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "EngUnit area cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "CATEGORIA_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });

            // Implement your logic to read user options from the specified file
            return new OpzioniUtente
            {
                FiltriPossibili = filtriPossibili
            };
        }

        public static bool IsBudgetFileOk(string filePath)
        {
            // check file existence

            // check expected worksheet names

            // check expected columns in each worksheet

            // Implement your logic to check if the budget file is OK
            return true;
        }
    }
}