using System;
using System.Collections.Generic;

namespace ReportRefresher.Entities
{
    public class UpdateReportsInput
    {
        #region Input obbligatorio
        public DateTime DataAggiornamento { get; private set; }
        public string FileController_FilePath { get; private set; }
        public string FileReport_FilePath { get; private set; }
        public string NewReport_FilePath { get; private set; }
        #endregion

        #region Input con valori di default
        public bool EvidenziaErroriNelFileDiInput { get; private set; }
        public string FileDebug_FilePath { get; private set; }
        #endregion

        #region Input opzionale
        public List<FornitoreCensito> FornitoriDaAggiungere { get; private set; }
        #endregion

        public int Periodo { get { return DataAggiornamento.Month; } }

        public int AnnoCorrente
        {
            get
            {
                if (DataAggiornamento.Month == 1)
                {
                    // A gennaio si elaborano i dati dell'anno precedente
                    return DataAggiornamento.Year - 1;
                }
                else
                {
                    return DataAggiornamento.Year;
                }
            }
        }



        public UpdateReportsInput(
            DateTime dataAggiornamento,
            string fileController_FilePath,
            string fileReport_FilePath,
            string newReport_FilePath,
            string fileDebug_FilePath = null,
            bool evidenziaErroriNelFileDiInput = false)
        {
            //if (string.IsNullOrWhiteSpace(fileController_FilePath))
            //    throw new ArgumentNullException(nameof(fileController_FilePath));
            //if (string.IsNullOrWhiteSpace(fileReport_FilePath))
            //    throw new ArgumentNullException(nameof(fileReport_FilePath));
            //if (string.IsNullOrWhiteSpace(newReport_FilePath))
            //    throw new ArgumentNullException(nameof(newReport_FilePath));

            DataAggiornamento = dataAggiornamento;
            FileController_FilePath = fileController_FilePath;
            FileReport_FilePath = fileReport_FilePath;
            NewReport_FilePath = newReport_FilePath;
            FileDebug_FilePath = fileDebug_FilePath;
            EvidenziaErroriNelFileDiInput = evidenziaErroriNelFileDiInput;
        }

        public void SettaNuoviFornitori(List<FornitoreCensito> fornitoriDaAggiungere)
        {
            FornitoriDaAggiungere = fornitoriDaAggiungere;
        }
    }
}