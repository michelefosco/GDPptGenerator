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
        public static GetUserOptionsFromDataSourceOutput GetUserOptionsFromDataSource(GetUserOptionsFromDataSourceInput getUserOptionsFromDataSourceInput)
        {
            var dataSourceTemplateFile = Path.Combine(getUserOptionsFromDataSourceInput.TemplatesFolder, Constants.FileNames.DATA_SOURCE_FILENAME);

         
            var filtriPossibili = new List<FilterItems>();
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "ProjType Cluster 2_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BussinessArea Cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BusinessArea_", ValoriSelezionati = new List<string> { "Valore 1", "Valore 2", "Valore 3", "Valore 4" } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "ProjType_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Proj type cluster 2", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Business", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "EngUnit area cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "CATEGORIA_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });

            return new GetUserOptionsFromDataSourceOutput
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
