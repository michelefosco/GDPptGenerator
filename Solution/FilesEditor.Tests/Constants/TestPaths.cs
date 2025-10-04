using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesEditor.Tests.Constants
{
    internal class TestPaths
    {
        #region Output paths
        public const string TEMPLATES_FOLDER = "Templates";

        public const string OUTPUT_FOLDER = "";
        public const string OUTPUT_FILE = "SoldReport.xlsx";
        public const string OUTPUT_DEBUGFILE = "SoldReport_Debugfile.xlsx";
        #endregion


        #region Input paths
        // Database folders
        public const string Database_OK_DatiAdHoc = @"Database\OK\DatiAdHoc";
        public const string Database_OK_DaGD = @"Database\OK\DaGD";
        public const string KO_MacchineNomiDuplicatiStessoFile = @"Database\KO\MacchineNomiDuplicatiStessoFile";
        public const string KO_MacchineNomiDuplicatiFileDiversi = @"Database\KO\MacchineNomiDuplicatiFileDiversi";


        // Venduto
        public const string Venduto_DatiAdHoc = @"Venduto\OK\DatiAdHoc";
        public const string Venduto_DatiDaGD = @"Venduto\OK\DaGD";

        public const string Venduto_TrovateUnknownMachines_01 = @"Venduto\KO\TrovateMacchineSconosciute_01";

        #endregion
    }
}
