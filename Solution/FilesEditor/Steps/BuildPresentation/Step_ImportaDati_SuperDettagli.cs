using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using OfficeOpenXml;
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

        private void AggiungiRigheInFondoAllaTabella(ExcelWorksheet destWorksheet, int numberOfRowsToBeAdded)
        {
            int NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA = 10000;

            // Numero iniziale di righe previsto dopo la cancellazione
            var numeroRighePrevistoDopoInserimento = destWorksheet.Dimension.End.Row + numberOfRowsToBeAdded;

            // Numero di righe ancora da cancellare
            var stillToBeAdded = numberOfRowsToBeAdded;
            var numberOfExceptionsToIgnore = 5 + (numberOfRowsToBeAdded / NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA);
            while (stillToBeAdded > 0)
            {
                var numberOfRowsToBeAddedInThisLoop = (stillToBeAdded > NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA)
                        ? NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA
                        : stillToBeAdded;

                try
                {
                    var startTime = DateTime.UtcNow;
                    destWorksheet.InsertRow(destWorksheet.Dimension.End.Row, numberOfRowsToBeAddedInThisLoop);
                    Context.DebugInfoLogger.LogPerformance(StepName + $" destWorksheet.InsertRow({destWorksheet.Dimension.End.Row}, {numberOfRowsToBeAddedInThisLoop});", DateTime.UtcNow - startTime);
                }
                catch (Exception ex)
                {
                    numberOfExceptionsToIgnore--;
                    if (numberOfExceptionsToIgnore >= 1)
                    {
                        if (NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA >= 2000)
                        {
                            NUMERO_MASSIMO_RIGHE_DA_AGGIUNGERE_PER_VOLTA -= 1000;
                        }

                        // Riprovo dopo una pausa
                        Thread.Sleep(1000);
                        continue;
                    }

                    // ho finito il numero di tentativi
                    throw new Exception($"Unable to add {numberOfRowsToBeAdded} rows to the 'Superdettagli' data sheet after multiple attempts.\n{ex.Message}");

                }

                stillToBeAdded = numeroRighePrevistoDopoInserimento - destWorksheet.Dimension.End.Row;
            }
        }

        private void CancellaBloccoDiRighe(ExcelWorksheet destWorksheet, int indexFirstRowOfTheBlock, int numberOfRowsToBeDeleted, string rowBelogingToSpecifiPeriodYear = null)
        {
            //var rangeToDelete = destWorksheet.Cells[
            //        indexFirstRowOfTheBlock,                                                // row start
            //        1,    // col start 
            //        indexFirstRowOfTheBlock+numberOfRowsToBeDeleted-1,                                               // row end
            //        destWorksheet.Dimension.End.Column // col end
            //        ];
            //rangeToDelete.Value = null;
            //rangeToDelete.Clear();

            const int NUMERO_MASSIMO_ASSOLUTO_DI_RIGHE_CANCELLABILI_PER_VOLTA = 8192;
            const int NUMERO_MASSIMO_DI_RIGHE_CANCELLABILI_IN_CASO_DI_ECCEZIONE = 4096;

            int dimensioneMassimaBlocchiDaCancellare = NUMERO_MASSIMO_ASSOLUTO_DI_RIGHE_CANCELLABILI_PER_VOLTA;
            var dimensioneBloccoDaUsarsiInCasoDiEccezione = NUMERO_MASSIMO_DI_RIGHE_CANCELLABILI_IN_CASO_DI_ECCEZIONE;


            // Numero iniziale di righe previsto dopo la cancellazione
            var numeroRighePrevistoDopoLaCancellazione = destWorksheet.Dimension.End.Row - numberOfRowsToBeDeleted;
            var numeroRigheIniziali = destWorksheet.Dimension.End.Row;

            bool eccezioneAvvenuta = false;

            // Numero di righe ancora da cancellare
            var stillToBeDeleted = numberOfRowsToBeDeleted;
            var numberOfExceptionsToIgnore = 200;
            var numeroRigheCancellatoFinora = 0;
            while (stillToBeDeleted > 0)
            {
                var numberOfRowsToBeDeletedInThisLoop = (stillToBeDeleted > dimensioneMassimaBlocchiDaCancellare)
                                        ? dimensioneMassimaBlocchiDaCancellare
                                        : stillToBeDeleted;

                // Cancello prima l'ultima parte del blocco in modo da avere meno shift delle righe
                var indiceFirstRowToBeDeletedInThisLoop = indexFirstRowOfTheBlock + stillToBeDeleted - numberOfRowsToBeDeletedInThisLoop;

                DateTime startTime = DateTime.UtcNow;
                try
                {
                    // inizio il conteggio del tempo per la cancellazione
                    startTime = DateTime.UtcNow;
                    Context.DebugInfoLogger.LogPerformance(StepName + $"- CancellaBloccoDiRighe() -  Provo destWorksheet.DeleteRow({indiceFirstRowToBeDeletedInThisLoop}, {numberOfRowsToBeDeletedInThisLoop});", DateTime.UtcNow - startTime);
                    destWorksheet.DeleteRow(indiceFirstRowToBeDeletedInThisLoop, numberOfRowsToBeDeletedInThisLoop);// Cancellaziione della parte in basso del blocco
                    //destWorksheet.DeleteRow(indexFirstRowOfTheBlock, numberOfRowsToBeDeletedInThisLoop); // Cancellaziione della parte in alto del blocco
                    Context.DebugInfoLogger.LogPerformance(StepName + $"- CancellaBloccoDiRighe() - OK;", DateTime.UtcNow - startTime);

                    // rivedo la dimensione massima del blocco da cancellare
                    if (eccezioneAvvenuta)
                    {
                        // riporto la dimensione massima del blocco da cancellare ad un valore più alto
                        dimensioneMassimaBlocchiDaCancellare = NUMERO_MASSIMO_DI_RIGHE_CANCELLABILI_IN_CASO_DI_ECCEZIONE;
                    }
                }
                catch (Exception ex)
                {
                    if (numberOfExceptionsToIgnore <= 0)
                    {
                        // Ho finito il numero di eccezioni da ignorare, sollevo l'eccezione
                        // preparo il messaggio per l'utente con o senza il riferimento all'anno specifico
                        var userMEssage = string.IsNullOrEmpty(rowBelogingToSpecifiPeriodYear)
                                ? $"Unable to delete {numberOfRowsToBeDeleted} rows from the “Superdettagli” data sheet.\r\nPlease open the DataSource file by clicking the “Data Source” menu and manually delete as many as possible deleteble rows."
                                : $"Unable to delete rows beloging to the year {rowBelogingToSpecifiPeriodYear} from the “Superdettagli” data sheet.\r\nPlease open the DataSource file by clicking on the “DataSource” menu and manually delete the rows belonging to the year {rowBelogingToSpecifiPeriodYear}.";
                        throw new ManagedException(
                            filePath: Context.DataSourceFilePath,
                            fileType: FileTypes.DataSource,
                            worksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
                            cellRow: null,
                            cellColumn: null,
                            valueHeader: ValueHeaders.None,
                            value: null,
                            errorType: ErrorTypes.UnhandledException,
                            userMessage: userMEssage);
                    }
                    else
                    {
                        // ignoro l'eccezione e riprovo
                        numberOfExceptionsToIgnore--;
                        Context.DebugInfoLogger.LogPerformance(StepName + $"- CancellaBloccoDiRighe() - Exception (rimangono ancora {numberOfExceptionsToIgnore} tentativi) durante destWorksheet.DeleteRow({indiceFirstRowToBeDeletedInThisLoop}, {numberOfRowsToBeDeletedInThisLoop});", DateTime.UtcNow - startTime);
                    }

                    // Gestisco il caso specifico dell'eccezione "Index was outside the bounds of the array."
                    if (ex.Message.Equals("Index was outside the bounds of the array.", StringComparison.Ordinal))
                    {
                        eccezioneAvvenuta = true;
                        // rivedo la dimensione massima del blocco da cancellare
                        dimensioneMassimaBlocchiDaCancellare = _getNuovaDimensioneMassimaDaUsare();

                        Context.DebugInfoLogger.LogPerformance(StepName + $"- CancellaBloccoDiRighe() - Imposto la dimensione dei blocchi da cancellare a {dimensioneMassimaBlocchiDaCancellare}", DateTime.UtcNow - startTime);

                        // Riprovo dopo una pausa
                        Thread.Sleep(500);
                    }
                }
                finally
                {
                    // ricarico il numero di righe cancellate e ancora da cancellare, ci potrebbe essere stato una parziale cancellazione delle righe
                    stillToBeDeleted = destWorksheet.Dimension.End.Row - numeroRighePrevistoDopoLaCancellazione;
                    numeroRigheCancellatoFinora = numeroRigheIniziali - destWorksheet.Dimension.End.Row;
                    Context.DebugInfoLogger.LogPerformance(StepName + $"- CancellaBloccoDiRighe() - Situazione: stillToBeDeleted={stillToBeDeleted} numeroRigheCancellatoFinora={numeroRigheCancellatoFinora}", DateTime.UtcNow - startTime);
                }
            }


            int _getNuovaDimensioneMassimaDaUsare()
            {
                // ripristina numero di dimensione massima
                if (dimensioneBloccoDaUsarsiInCasoDiEccezione > 1)
                {
                    dimensioneMassimaBlocchiDaCancellare = dimensioneBloccoDaUsarsiInCasoDiEccezione;
                    // dimezzo il valore per il prossimo tentativo
                    dimensioneBloccoDaUsarsiInCasoDiEccezione /= 2;
                }

                return dimensioneMassimaBlocchiDaCancellare;
            }
        }



        private void CancellaBloccoDiRighe_Bis(ExcelWorksheet destWorksheet, int indexFirstRowOfTheBlock, int numberOfRowsToBeDeleted)
        {
            // Numero iniziale di righe previsto dopo la cancellazione
            var numeroRighePrevistoDopoLaCancellazione = destWorksheet.Dimension.End.Row - numberOfRowsToBeDeleted;

            // Numero di righe ancora da cancellare
            var stillToBeDeleted = numberOfRowsToBeDeleted;
            var numberOfExceptionsToIgnore = 10;

            while (stillToBeDeleted > 0)
            {
                try
                {
                    var startTime = DateTime.UtcNow;
                    destWorksheet.DeleteRow(indexFirstRowOfTheBlock, stillToBeDeleted);
                    Context.DebugInfoLogger.LogPerformance(StepName + $" destWorksheet.DeleteRow({indexFirstRowOfTheBlock}, {stillToBeDeleted});", DateTime.UtcNow - startTime);
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