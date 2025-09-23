using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Ordinamento tabelle
    /// </summary>
    internal class Step_OrdinamentoTabelle : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            OrdinaTabella_AnagraficaFornitori(context.InfoFileReport, context.Configurazione);
            OrdinaTabella_ListaDati(context.InfoFileReport, context.Configurazione);
            OrdinaTabella_BudgetStudiIpotesi1(context.InfoFileReport, context.Configurazione);
            OrdinaTabella_Reportistica(context.InfoFileReport, context.Configurazione);

            return null;
        }

        private void OrdinaTabella_AnagraficaFornitori(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_AnagraficaFornitori; // Anagrafica fornitori

            var ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.AnagraficaFornitori_PrimaRigaFornitori, configurazione.AnagraficaFornitori_ColonnaSigle);
            VerificaAssenzaFormuleNellaColonnaSigle(
                    infoFileReport: infoFileReport,
                    worksheetName: worksheetName,
                    colToBeChecked: configurazione.AnagraficaFornitori_ColonnaSigle,
                    rowFrom: configurazione.AnagraficaFornitori_PrimaRigaFornitori,
                    rowTo: ultimarigaUtilizzata
                    );

            if (ultimarigaUtilizzata <= configurazione.AnagraficaFornitori_PrimaRigaFornitori)
            { return; /* meno di 2 righe presenti, non è neccessario ordinare*/}

            infoFileReport.EPPlusHelper.OrdinaTabella(
                worksheetName: worksheetName,
                fromRow: configurazione.AnagraficaFornitori_PrimaRigaFornitori,
                fromCol: configurazione.AnagraficaFornitori_ColonnaSigle,
                toRow: ultimarigaUtilizzata,
                toCol: configurazione.AnagraficaFornitori_ColonnaNomi
                );
        }

        private void OrdinaTabella_BudgetStudiIpotesi1(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_BudgetStudiIpotesi; // BUDGET STUDI -  IPOTESI 1

            var ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.IpotesiStudio_PrimaRigaFornitori, configurazione.IpotesiStudio_ColonnaSigle);
            if (ultimarigaUtilizzata <= configurazione.IpotesiStudio_PrimaRigaFornitori)
            { return; /* meno di 2 righe presenti, non è neccessario ordinare*/}


            VerificaAssenzaFormuleNellaColonnaSigle(
                infoFileReport: infoFileReport,
                worksheetName: worksheetName,
                colToBeChecked: configurazione.IpotesiStudio_ColonnaSigle,
                rowFrom: configurazione.IpotesiStudio_PrimaRigaFornitori,
                rowTo: ultimarigaUtilizzata
                );

            infoFileReport.EPPlusHelper.OrdinaTabella(
                worksheetName: worksheetName,
                fromRow: configurazione.IpotesiStudio_PrimaRigaFornitori,
                fromCol: configurazione.IpotesiStudio_ColonnaSigle,
                toRow: ultimarigaUtilizzata,
                toCol: configurazione.IpotesiStudio_ColonnaUltimaDaOrdinare
                );
        }
        private void OrdinaTabella_Reportistica(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_Reportistica; // REPORTISTICA

            var ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.Reportistica_PrimaRigaFornitori, configurazione.Reportistica_ColonnaSigle_SX);
            if (ultimarigaUtilizzata <= configurazione.Reportistica_PrimaRigaFornitori)
            { return; /* meno di 2 righe presenti, non è neccessario ordinare*/}

            VerificaAssenzaFormuleNellaColonnaSigle(
                infoFileReport: infoFileReport,
                worksheetName: worksheetName,
                colToBeChecked: configurazione.Reportistica_ColonnaSigle_SX,
                rowFrom: configurazione.Reportistica_PrimaRigaFornitori,
                rowTo: ultimarigaUtilizzata
                );

            // Ordina la tabella SX
            infoFileReport.EPPlusHelper.OrdinaTabella(
                 worksheetName: worksheetName,
                 fromRow: configurazione.Reportistica_PrimaRigaFornitori,
                 fromCol: configurazione.Reportistica_ColonnaSigle_SX,
                 toRow: ultimarigaUtilizzata,
                 toCol: configurazione.Reportistica_ColonnaSigle_SX
                 );

            VerificaAssenzaFormuleNellaColonnaSigle(
                infoFileReport: infoFileReport,
                worksheetName: worksheetName,
                colToBeChecked: configurazione.Reportistica_ColonnaSigle_DX,
                rowFrom: configurazione.Reportistica_PrimaRigaFornitori,
                rowTo: ultimarigaUtilizzata
                );

            ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.Reportistica_PrimaRigaFornitori, configurazione.Reportistica_ColonnaSigle_DX);
            // Ordina la tabella DX
            infoFileReport.EPPlusHelper.OrdinaTabella(
                 worksheetName: worksheetName,
                 fromRow: configurazione.Reportistica_PrimaRigaFornitori,
                 fromCol: configurazione.Reportistica_ColonnaSigle_DX,
                 toRow: ultimarigaUtilizzata,
                 toCol: configurazione.Reportistica_ColonnaSigle_DX
                 );
        }

    }
}