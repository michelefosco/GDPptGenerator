using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using OfficeOpenXml;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_SuperDettagli : StepBase
    {
        internal override string StepName => "Step_ImportaDati_SuperDettagli";

        public Step_ImportaDati_SuperDettagli(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ImportaFile_Superdettagli();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportaFile_Superdettagli()
        {
            var sourceFilePath = Context.FileSuperDettagliPath;
            var sourceWorksheetName = WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA;
            var sourceHeadersRow = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW;
            int sourceHeadersFirstColumn = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL;
            //
            var destWorksheetName = WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA;
            var destHeadersRow = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW;
            var destHeadersFirstColumn = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL;
            //
            var valoriColonnaAnnoRigheDaValutare = new List<string>();
            var periodYearString = Context.PeriodYear.ToString();
            var infoRowsDestinazione = new InfoRows();  // Variabili per il conteggio delle righe elaborate


            #region Lettura degli headers del foglio sorgente e foglio destinazione
            var destHeaders = Context.DataSourceEPPlusHelper.GetHeadersFromRow(destWorksheetName, destHeadersRow, destHeadersFirstColumn, true).Select(_ => _.ToLower()).ToList();
            var numeroDiColonneDaCopiare = destHeaders.Count;
            #endregion


            #region WorkSheets sorgente e destinazione
            // Foglio sorgente
            var sourceWorksheet = Context.SuperdettagliFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[sourceWorksheetName];

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            infoRowsDestinazione.Iniziali = destWorksheet.Dimension.Rows - destHeadersRow;
            destWorksheet.Select(destWorksheet.Cells[1, 1]);
            #endregion


            #region Applico gli eventuali filtri al file sorgente
            ApplicaFiltriAllaSorgenteDati(sourceWorksheet);

            // righe presenti nel file sorce dopo l'applicazione dei filtri
            var numeroRigheDaImportareDaSorgente = sourceWorksheet.Dimension.End.Row - sourceHeadersRow;
            if (numeroRigheDaImportareDaSorgente == 0)
            {
                Context.AddWarning($"The file '{FileTypes.SuperDettagli}' contains no rows that meet the applied filters.");
                return; // interrompo l'esecuzione
            }
            #endregion


            #region Calcolo il numero di righe da riusare o cancellare nella tabella di destinazione
            int numeroRigheDestinazione_DaRiusareOppureCancellare = 0;
            if (Context.AppendCurrentYear_FileSuperDettagli)
            {
                valoriColonnaAnnoRigheDaValutare = GetListaValoriAnno(destWorksheet);
                numeroRigheDestinazione_DaRiusareOppureCancellare = valoriColonnaAnnoRigheDaValutare.Count(_ => string.Equals(_, periodYearString, StringComparison.Ordinal));
            }
            else
            {
                numeroRigheDestinazione_DaRiusareOppureCancellare = infoRowsDestinazione.Iniziali;
            }
            infoRowsDestinazione.Preservate = infoRowsDestinazione.Iniziali - numeroRigheDestinazione_DaRiusareOppureCancellare;
            #endregion


            #region Aggiungo o elimino righe nella tabella di destinazione per fare spazio alle righe da importare (o evenutalmente rimuovere quelle non più necessarie in quanto ho meno righe da importare di quelle da sostituire)
            if (numeroRigheDaImportareDaSorgente == numeroRigheDestinazione_DaRiusareOppureCancellare)
            {
                // nessuna riga da aggiungere o eliminare, sostituisco tutte le righe attuali con le nuove
                infoRowsDestinazione.Aggiunte = 0;
                infoRowsDestinazione.Riutilizzate = numeroRigheDestinazione_DaRiusareOppureCancellare;
                infoRowsDestinazione.Eliminate = 0;
            }

            if (numeroRigheDaImportareDaSorgente > numeroRigheDestinazione_DaRiusareOppureCancellare)
            {
                // Indice dell'ultima riga con dati prima dell'inserimento delle nuove righe. Qui sposterò (ripristinando la loro posizione) i dati dell'ultima riga che sarà stata spostata in basso per fare spazio alle nuove righe aggiungente
                var indiceUltimaRigaConDatiPrimaDellaImportazione = destWorksheet.Dimension.End.Row;

                //numeroRigheRiutilizzate = 0;
                // Aggiungo delle righe a quelle "riutilizzate" per ospitare tutte le righe da importare
                var numeroRigheDaAggiungere = numeroRigheDaImportareDaSorgente - numeroRigheDestinazione_DaRiusareOppureCancellare;
                AggiungiRigheInFondoAllaTabella(destWorksheet, numeroRigheDaAggiungere);
                infoRowsDestinazione.Aggiunte = numeroRigheDaAggiungere;
                infoRowsDestinazione.Riutilizzate = numeroRigheDestinazione_DaRiusareOppureCancellare;
                infoRowsDestinazione.Eliminate = 0;

                //aggiungo anche elementi vuoti alla lista in modo da rispecchiare l'allineamento con le righe del foglio di destinazione
                valoriColonnaAnnoRigheDaValutare.AddRange(Enumerable.Repeat(string.Empty, numeroRigheDaAggiungere));

                if (Context.AppendCurrentYear_FileSuperDettagli)
                {
                    // Sposto il contenuto dell'ultima riga nella sua posizione originale. Qesto è necessario solo in caso di appendo su Superdettagli per preservare il contenuto dei dati
                    SpostaRiga(destWorksheet, destWorksheet.Dimension.End.Row, indiceUltimaRigaConDatiPrimaDellaImportazione, numeroDiColonneDaCopiare);
                }
            }

            if (numeroRigheDaImportareDaSorgente < numeroRigheDestinazione_DaRiusareOppureCancellare)
            {
                // Le righe "DaRiusareOppureCancellare" sono più di quellle necessarie ai dati da importare. Non le userò tutte. Alcune verranno cancellate
                var numeroRigheDaEliminare = numeroRigheDestinazione_DaRiusareOppureCancellare - numeroRigheDaImportareDaSorgente;

                if (Context.AppendCurrentYear_FileSuperDettagli)
                {
                    // cancellazione di <numeroRigheDaEliminare> righe a partire dalla fine di solo quelle con anno corrente)
                    CancellaTotRigheConAnnoCorrenteDalFondoDellaTabella(destWorksheet, numeroRigheDaEliminare, valoriColonnaAnnoRigheDaValutare, periodYearString, destHeadersRow);

                    // devo rivalutare i valori della colonna anno dopo la cancellazione in quanto potrebbero essere cambiate le posizioni intermedie e non solo eliminate le ultime
                    valoriColonnaAnnoRigheDaValutare = GetListaValoriAnno(destWorksheet);
                }
                else
                {
                    // cancellazione incodizionata di <numeroRigheDaEliminare> righe a partire dalla fine
                    // rimarrà almeno una riga in quanto il numero di righe da importare è maggiore di zero
                    CancellaBloccoDiRighe(destWorksheet, destHeadersRow + 1, numeroRigheDaEliminare);
                }
                infoRowsDestinazione.Riutilizzate = numeroRigheDestinazione_DaRiusareOppureCancellare - numeroRigheDaEliminare;
                infoRowsDestinazione.Eliminate = numeroRigheDaEliminare;
                infoRowsDestinazione.Aggiunte = 0;
            }
            #endregion


            #region Copiare le righe da importare nei blocchi di righe con anno == Context.PeriodYear (se in modalità append su Superdettagli), indistintamente altrimenti
            // Deterimo i blocchi da riempire nella tabella di destinazione
            var blocchiDaRiempire = new List<InfoBloccoDaCopiare>();

            if (Context.AppendCurrentYear_FileSuperDettagli)
            {
                blocchiDaRiempire = GetBlocchiRigheDaRiempire(valoriColonnaAnnoRigheDaValutare, periodYearString);
            }
            else
            {
                // A questo punto il numero di righe da importare è esattamente uguale al numero di righe presenti nella tabella di destinazione
                // Copio tutte le righe indiscriminatamente in un unico blocco
                blocchiDaRiempire.Add(new InfoBloccoDaCopiare
                {
                    IndicePrimaRigaSorgente = sourceHeadersRow + 1,
                    IndicePrimaRigaDestinazione = destHeadersRow + 1,
                    NumeroRigheNelBlocco = sourceWorksheet.Dimension.Rows - sourceHeadersRow
                });
            }



            foreach (var blocco in blocchiDaRiempire)
            {
                CopiaBloccoRigheDaSorgenteToDestinazione(sourceWorksheet: sourceWorksheet,
                                                        indicePrimaRigaSorgente: blocco.IndicePrimaRigaSorgente,
                                                        numberRigheDaCopiare: blocco.NumeroRigheNelBlocco,
                                                        numeroDiColonneDaCopiare: numeroDiColonneDaCopiare,
                                                        destWorksheet: destWorksheet,
                                                        indicePrimaRigaDestinazione: blocco.IndicePrimaRigaDestinazione
                                                        );
            }

            var numeroRigheNeiBlocchi = blocchiDaRiempire.Sum(_ => _.NumeroRigheNelBlocco);
            if (numeroRigheNeiBlocchi != numeroRigheDaImportareDaSorgente)
            {
                throw new Exception($"The number of rows copied ({numeroRigheNeiBlocchi}) does not match the number of rows to import ({numeroRigheDaImportareDaSorgente}).");
            }
            #endregion


            #region Log delle informazioni sulle righe
            infoRowsDestinazione.Finali = destWorksheet.Dimension.End.Row - destHeadersRow;
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.SuperDettagli, infoRowsDestinazione);
            #endregion


            #region Just for debug: verifico che le righe copiate siano corrette
            infoRowsDestinazione.VerificaCoerenzaValori();
            var numeroRigheConAnnoCorrente = GetListaValoriAnno(destWorksheet).Count(_ => string.Equals(_, periodYearString, StringComparison.Ordinal));
            if (numeroRigheConAnnoCorrente != numeroRigheDaImportareDaSorgente)
            {
                throw new Exception($"The number of rows with year '{periodYearString}' in the destination sheet ({numeroRigheConAnnoCorrente}) does not match the number of rows imported from source ({numeroRigheDaImportareDaSorgente}).");
            }
            #endregion

            Context.SuperdettagliFileEPPlusHelper.Close();
        }

        private List<InfoBloccoDaCopiare> GetBlocchiRigheDaRiempire(List<string> valoriColonnaAnnoRigheDaValutare, string periodYearString)
        {
            var sourceHeadersRow = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW;
            var destHeadersRow = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW;

            var blocchiDaRiempire = new List<InfoBloccoDaCopiare>();

            int indicePrimaRigaSorgente = sourceHeadersRow + 1;
            int IndicePrimaRigaDestinazione = -1;
            int NumeroRigheNelBlocco = 0;

            int conteggioNumeroCopiate = 0;


            for (var soruceRowsOffsetZeroBased = 0; soruceRowsOffsetZeroBased <= valoriColonnaAnnoRigheDaValutare.Count - 1; soruceRowsOffsetZeroBased++)
            {
                // Individuo le righe il cui valore nella cella "anno" è uguale a Context.PeriodYear
                var valoreAnnoRigaCorrente = valoriColonnaAnnoRigheDaValutare[soruceRowsOffsetZeroBased];
                if (string.Equals(valoreAnnoRigaCorrente, periodYearString, StringComparison.Ordinal) || string.IsNullOrEmpty(valoreAnnoRigaCorrente))
                {
                    // La riga corrente appartitente ad un blocco da copiare determino se la riga inizia un nuovo blocco o continua uno già iniziato
                    if (IndicePrimaRigaDestinazione == -1)
                    {
                        // ho individuato un nuovo blocco di righe da riempire
                        IndicePrimaRigaDestinazione = destHeadersRow + 1 + soruceRowsOffsetZeroBased;
                        NumeroRigheNelBlocco = 1;
                    }
                    else
                    {
                        // il blocco precedentemente individuato continua
                        NumeroRigheNelBlocco++;
                    }
                }
                else
                {
                    // questa riga non appartienre da un blocco, determino se ha interroto un blocco già iniziato
                    if (NumeroRigheNelBlocco > 0)
                    {
                        // c'era un blocco in corso, lo aggiungo
                        blocchiDaRiempire.Add(new InfoBloccoDaCopiare
                        {
                            IndicePrimaRigaSorgente = indicePrimaRigaSorgente,
                            IndicePrimaRigaDestinazione = IndicePrimaRigaDestinazione,
                            NumeroRigheNelBlocco = NumeroRigheNelBlocco
                        });
                        indicePrimaRigaSorgente += NumeroRigheNelBlocco;


                        conteggioNumeroCopiate += NumeroRigheNelBlocco;
                    }

                    // resetto gli indici del blocco
                    NumeroRigheNelBlocco = 0;
                    IndicePrimaRigaDestinazione = -1;
                }
            }

            // verifico alla fine dei dati c'èera un blocco in corso
            if (NumeroRigheNelBlocco > 0)
            {
                // c'era un blocco in corso, lo aggiungo
                blocchiDaRiempire.Add(new InfoBloccoDaCopiare
                {
                    IndicePrimaRigaSorgente = indicePrimaRigaSorgente,
                    IndicePrimaRigaDestinazione = IndicePrimaRigaDestinazione,
                    NumeroRigheNelBlocco = NumeroRigheNelBlocco
                });
            }

            return blocchiDaRiempire;
        }

        private void CopiaBloccoRigheDaSorgenteToDestinazione(ExcelWorksheet sourceWorksheet, int indicePrimaRigaSorgente, int numberRigheDaCopiare, int numeroDiColonneDaCopiare,
                                                                ExcelWorksheet destWorksheet, int indicePrimaRigaDestinazione)
        {
            var indiceUltimaRigaSorgente = indicePrimaRigaSorgente + numberRigheDaCopiare - 1;
            var indiceUltimaRigaDestinazione = indicePrimaRigaDestinazione + numberRigheDaCopiare - 1;

            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                                indicePrimaRigaSorgente,                                                // row start
                                Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,    // col start 
                                indiceUltimaRigaSorgente,                                               // row end
                                Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare - 1  // col end
                                ];

            var destRange = destWorksheet.Cells[
                                indicePrimaRigaDestinazione,                                        // row start
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,  // col start 
                                indiceUltimaRigaDestinazione,                                       // row end
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare - 1   // col end
                                ];
            destRange.Value = sourceRange.Value;


            var yearRange = destWorksheet.Cells[
                    indicePrimaRigaDestinazione,                                // row start
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL,   // col start 
                    indiceUltimaRigaDestinazione,                               // row end
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL    // col end
                    ];
            yearRange.Value = Context.PeriodYear;
        }

        private void CancellaTotRigheConAnnoCorrenteDalFondoDellaTabella(ExcelWorksheet destWorksheet, int numeroRigheDaEliminare, List<string> valoriColonnaAnnoRigheDaValutare, string periodYearString, int destHeadersRow)
        {
            var startTime = DateTime.UtcNow;

            int indexLastRowOfTheBlock = -1;
            int numberOfRowsInTheBlock = 0;

            var numeroRigheAncoraDaEliminare = numeroRigheDaEliminare;

            for (var valueIndex = valoriColonnaAnnoRigheDaValutare.Count - 1; valueIndex >= 0; valueIndex--)
            {
                if (numberOfRowsInTheBlock >= numeroRigheAncoraDaEliminare)
                {
                    // ho raggiunto il numero ri righe da eliminare, interrompo il blocco corrente anche se potenzialmente anora più grande
                    var indexFirstRowOfTheBlock = indexLastRowOfTheBlock - numberOfRowsInTheBlock + 1;
                    CancellaBloccoDiRighe(destWorksheet, indexFirstRowOfTheBlock, numberOfRowsInTheBlock);
                    return;
                }

                // elimino le righe il cui valore nella cella "anno" è uguale a Context.PeriodYear
                var valoreAnnoRigaCorrente = valoriColonnaAnnoRigheDaValutare[valueIndex];
                if (string.Equals(valoreAnnoRigaCorrente, periodYearString, StringComparison.Ordinal))
                {
                    if (indexLastRowOfTheBlock == -1)
                    {
                        indexLastRowOfTheBlock = valueIndex + 1 + destHeadersRow;
                        numberOfRowsInTheBlock = 1;
                    }
                    else
                    {
                        numberOfRowsInTheBlock++;
                    }
                }
                else
                {
                    if (numberOfRowsInTheBlock > 0)
                    {
                        var indexFirstRowOfTheBlock = indexLastRowOfTheBlock - numberOfRowsInTheBlock + 1;
                        CancellaBloccoDiRighe(destWorksheet, indexFirstRowOfTheBlock, numberOfRowsInTheBlock);
                        numeroRigheAncoraDaEliminare -= numberOfRowsInTheBlock;
                        numberOfRowsInTheBlock = 0;
                    }

                    indexLastRowOfTheBlock = -1;
                }
            }
            if (numberOfRowsInTheBlock > 0)
            {
                var indexFirstRowOfTheBlock = indexLastRowOfTheBlock - numberOfRowsInTheBlock + 1;
                CancellaBloccoDiRighe(destWorksheet, indexFirstRowOfTheBlock, numberOfRowsInTheBlock);
            }

            Context.DebugInfoLogger.LogPerformance(StepName + $" CancellaTotRigheConAnnoCorrenteDalFondoDellaTabella {numeroRigheDaEliminare}", DateTime.UtcNow - startTime);
        }

        private void SpostaRiga(ExcelWorksheet worksheet, int rigaFrom, int rigaTo, int numeroDiColonneDaCopiare)
        {
            var copyRange = worksheet.Cells[
                               rigaFrom,                                                        // row start
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,                            // col start 
                                rigaFrom, // row end
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare - 1  // col end
                                ];

            var pastRange = worksheet.Cells[
                                rigaTo,                                                        // row start
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,                              // col start 
                                rigaTo,                                                       // row end
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare - 1   // col end
                                ];
            pastRange.Value = copyRange.Value;
            copyRange.Value = null;


            // Copio separatamente la colonna dell'anno
            copyRange = worksheet.Cells[
                                rigaFrom,                                                   // row start
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL, // col "year"
                                rigaFrom, // row end
                                Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL // col "year"
                                ];
            pastRange = worksheet.Cells[
                               rigaTo,                                                        // row start
                               Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL, // col "year"
                               rigaTo,                                                       // row end
                               Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL // col "year"
                               ];
            pastRange.Value = copyRange.Value;
            copyRange.Value = null;
        }

        private void ApplicaFiltriAllaSorgenteDati(ExcelWorksheet sourceWorksheet)
        {
            //todo: applicare filtri usando gli array come per il campo anno
            #region Sfoltisco le righe della sorgente in base ai filtri impostati
            var activeFilters = Context.ApplicableFilters.Where(_ => _.Table == InputDataFilters_Tables.SUPERDETTAGLI && _.SelectedValues.Any()).ToList();
            if (activeFilters.Any())
            {
                var sourceWorksheetName = WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA;
                var souceHeadersRow = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW;
                int sourceHeadersFirstColumn = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL;
                var sourceHeaders = Context.SuperdettagliFileEPPlusHelper.GetHeadersFromRow(sourceWorksheetName, souceHeadersRow, sourceHeadersFirstColumn, true).Select(_ => _.ToLower()).ToList();

                foreach (var filter in activeFilters)
                {
                    if (!sourceHeaders.Any(_ => _.Equals(filter.FieldName, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        throw new ArgumentOutOfRangeException($"One of the filters set does not have a corresponding header in the source file");
                    }

                    // trovo la colonna sorgente usando il nome della colonna di destinazione
                    int sourceColumnIndex = sourceHeaders.IndexOf(filter.FieldName.ToLower()) + 1;

                    //for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= sourceWorksheet.Dimension.End.Row; rowSourceIndex++)
                    for (var rowSourceIndex = sourceWorksheet.Dimension.End.Row; rowSourceIndex >= souceHeadersRow + 1; rowSourceIndex--)
                    {
                        // prendo il valore dalla sorgente
                        var value = sourceWorksheet.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // se il valore (non null) non è presente tra i valori selezionati, la riga viene saltata
                        if (value != null && !filter.SelectedValues.Any(_ => _.Equals(value)))
                        {
                            sourceWorksheet.DeleteRow(rowSourceIndex, 1, true);
                            continue;
                        }
                    }
                }
            }
            #endregion
        }

        private void AggiungiRigheInFondoAllaTabella_BK(ExcelWorksheet destWorksheet, int numeroRigheDaAccodare)
        {
            var numeroRighePrevistoAllaFineDellaAccodamento = destWorksheet.Dimension.End.Row + numeroRigheDaAccodare;

            int numeroTentativiDisponibili = 5;
            while (numeroTentativiDisponibili >= 1)
            {
                try
                {
                    numeroTentativiDisponibili--;
                    destWorksheet.InsertRow(destWorksheet.Dimension.End.Row, numeroRigheDaAccodare);
                }
                catch (Exception ex)
                {
                    if (numeroTentativiDisponibili >= 1)
                    {
                        // Riprovo dopo una pausa
                        Thread.Sleep(1000);
                        continue;
                    }

                    // ho finito il numero di tentativi
                    throw new Exception($"Unable to append {numeroRigheDaAccodare} rows to the 'Superdettagli' data sheet after multiple attempts.\n{ex.Message}");

                }

                if (destWorksheet.Dimension.End.Row == numeroRighePrevistoAllaFineDellaAccodamento)
                {
                    // Obiettivo raggiunto
                    break;
                }
                numeroRigheDaAccodare = numeroRighePrevistoAllaFineDellaAccodamento - destWorksheet.Dimension.End.Row;
            }
        }
        private void AggiungiRigheInFondoAllaTabella(ExcelWorksheet destWorksheet, int numberOfRowsToBeAdded)
        {
            const int NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 10000;
            //int NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 10000;

            // Numero iniziale di righe previsto dopo la cancellazione
            var numeroRighePrevistoDopoInserimento = destWorksheet.Dimension.End.Row + numberOfRowsToBeAdded;

            // Numero di righe ancora da cancellare
            var toBeAdded = numberOfRowsToBeAdded;
            var numberOfExceptionsToIgnore = 5 + (numberOfRowsToBeAdded / NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA);
            while (toBeAdded > 0)
            {
                if (toBeAdded > NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA)
                {
                    toBeAdded = NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA;

                    //#region Logica di adattamento del numero di righe da aggiungere per volta
                    //if (NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA == 6000)
                    //{
                    //    NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 8000;
                    //}
                    //else if (NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA == 8000)
                    //{
                    //    NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 10000;
                    //}
                    //else if (NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA == 10000)
                    //{
                    //    NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 6000;
                    //}

                    //#endregion
                }

                try
                {
                    var startTime = DateTime.UtcNow;
                    destWorksheet.InsertRow(destWorksheet.Dimension.End.Row, toBeAdded);
                    Context.DebugInfoLogger.LogPerformance(StepName + $" destWorksheet.InsertRow({destWorksheet.Dimension.End.Row}, {toBeAdded});", DateTime.UtcNow - startTime);
                }
                catch (Exception ex)
                {
                    numberOfExceptionsToIgnore--;
                    if (numberOfExceptionsToIgnore >= 1)
                    {
                        // Riprovo dopo una pausa
                        Thread.Sleep(1000);
                        continue;
                    }

                    // ho finito il numero di tentativi
                    throw new Exception($"Unable to add {numberOfRowsToBeAdded} rows to the 'Superdettagli' data sheet after multiple attempts.\n{ex.Message}");

                }

                toBeAdded = numeroRighePrevistoDopoInserimento - destWorksheet.Dimension.End.Row;
            }
        }

        private void CancellaBloccoDiRighe(ExcelWorksheet destWorksheet, int indexFirstRowOfTheBlock, int numberOfRowsToBeDeleted)
        {
            const int NUMERO_MASSIMO_RIGHE_CANCELLABILI_PER_VOLTA = 10000;

            // Numero iniziale di righe previsto dopo la cancellazione
            var numeroRighePrevistoDopoLaCancellazione = destWorksheet.Dimension.End.Row - numberOfRowsToBeDeleted;

            // Numero di righe ancora da cancellare
            var stillToBeDeleted = numberOfRowsToBeDeleted;
            var numberOfExceptionsToIgnore = 5 + (numberOfRowsToBeDeleted / NUMERO_MASSIMO_RIGHE_CANCELLABILI_PER_VOLTA);
            while (stillToBeDeleted > 0)
            {
                var numberOfRowsToBeDeletedInThisLoop = (stillToBeDeleted > NUMERO_MASSIMO_RIGHE_CANCELLABILI_PER_VOLTA)
                                        ? NUMERO_MASSIMO_RIGHE_CANCELLABILI_PER_VOLTA
                                        : stillToBeDeleted;

                // Cancello prima l'ultima parte del blocco in modo da avere meno shift delle righe
                var indiceFirstRowToBeDeletedInThisLoop = indexFirstRowOfTheBlock + stillToBeDeleted - numberOfRowsToBeDeletedInThisLoop;

                try
                {
                    var startTime = DateTime.UtcNow;
                    destWorksheet.DeleteRow(indiceFirstRowToBeDeletedInThisLoop, numberOfRowsToBeDeletedInThisLoop);
                    Context.DebugInfoLogger.LogPerformance(StepName + $" destWorksheet.DeleteRow({indiceFirstRowToBeDeletedInThisLoop}, {numberOfRowsToBeDeletedInThisLoop});", DateTime.UtcNow - startTime);
                }
                catch (Exception ex)
                {
                    numberOfExceptionsToIgnore--;
                    if (numberOfExceptionsToIgnore >= 1)
                    {
                        // Riprovo dopo una pausa
                        Thread.Sleep(1000);
                        continue;
                    }

                    // ho finito il numero di tentativi
                    throw new Exception($"Unable to delete {numberOfRowsToBeDeleted} rows from the 'Superdettagli' data sheet after multiple attempts.\n{ex.Message}");

                }

                stillToBeDeleted = destWorksheet.Dimension.End.Row - numeroRighePrevistoDopoLaCancellazione;
            }
        }

        private void IncollaRange(ExcelRange sourceRange, ExcelRange destRange)
        {
            if (sourceRange.Rows != destRange.Rows)
            { throw new Exception("Dimensione range errata (Rows)"); }

            if (sourceRange.Columns != destRange.Columns)
            { throw new Exception("Dimensione range errata (Columns)"); }


            int numeroTentativiDisponibili = 5;
            while (numeroTentativiDisponibili >= 1)
            {
                try
                {
                    destRange.Value = sourceRange.Value;
                    return;
                }
                catch (Exception ex)
                {
                    numeroTentativiDisponibili--;
                    if (numeroTentativiDisponibili >= 1)
                    {
                        // Riprovo dopo una pausa
                        Thread.Sleep(1000);
                        continue;
                    }

                    // ho finito il numero di tentativi
                    throw new Exception($"Unable to paste a range o data from {sourceRange.Address} to {destRange.Address}.\n{ex.Message}");
                }
            }
        }


        private List<string> GetListaValoriAnno(ExcelWorksheet destWorksheet)
        {
            var primaRigaConDati = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW + 1;
            var ultimaRigaConDati = destWorksheet.Dimension.End.Row;
            return destWorksheet.Cells[
                    primaRigaConDati,
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL,
                    ultimaRigaConDati,
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_YEAR_COL]
            .Select(c => c.Text).ToList();
        }

        private struct InfoBloccoDaCopiare
        {
            public int IndicePrimaRigaSorgente;
            public int NumeroRigheNelBlocco;
            public int IndicePrimaRigaDestinazione;
        }
    }
}