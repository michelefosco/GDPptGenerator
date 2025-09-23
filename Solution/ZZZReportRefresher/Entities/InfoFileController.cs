using EPPlusExtensions;
using System;

namespace ReportRefresher.Entities
{
    public class InfoFileController : InfoFileBase
    {
        public readonly string WorksheetName_Recap;
        public readonly string WorksheetName_ActualSoloCdc;
        public readonly string WorksheetName_FBL3Nact;
        public readonly string WorksheetName_CommitmentSoloCdc;
        public readonly string WorksheetName_ME5A;
        public readonly string WorksheetName_ME2N;

        public InfoFileController(EPPlusHelper epPlusHelper, string filePath,
            string worksheetName_Recap,
            string worksheetName_ActualSoloCdc,
            string worksheetName_FBL3Nact,
            string worksheetName_CommitmentSoloCdc,
            string worksheetName_ME5A,
            string worksheetName_ME2N) : base(epPlusHelper, filePath)
        {
            if (string.IsNullOrWhiteSpace(worksheetName_Recap))
                throw new ArgumentNullException(nameof(worksheetName_Recap));
            if (string.IsNullOrWhiteSpace(worksheetName_ActualSoloCdc))
                throw new ArgumentNullException(nameof(worksheetName_ActualSoloCdc));
            if (string.IsNullOrWhiteSpace(worksheetName_FBL3Nact))
                throw new ArgumentNullException(nameof(worksheetName_FBL3Nact));
            if (string.IsNullOrWhiteSpace(worksheetName_CommitmentSoloCdc))
                throw new ArgumentNullException(nameof(worksheetName_CommitmentSoloCdc));
            if (string.IsNullOrWhiteSpace(worksheetName_ME5A))
                throw new ArgumentNullException(nameof(worksheetName_ME5A));
            if (string.IsNullOrWhiteSpace(worksheetName_ME2N))
                throw new ArgumentNullException(nameof(worksheetName_ME2N));

            WorksheetName_Recap = worksheetName_Recap;
            WorksheetName_ActualSoloCdc = worksheetName_ActualSoloCdc;
            WorksheetName_FBL3Nact = worksheetName_FBL3Nact;
            WorksheetName_CommitmentSoloCdc = worksheetName_CommitmentSoloCdc;
            WorksheetName_ME5A = worksheetName_ME5A;
            WorksheetName_ME2N = worksheetName_ME2N;
        }
    }
}