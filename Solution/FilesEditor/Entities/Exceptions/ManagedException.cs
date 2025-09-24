using FilesEditor.Enums;
using System;

namespace FilesEditor.Entities.Exceptions
{
    public class ManagedException : Exception
    {
        public readonly TipologiaErrori TipologiaErrore;
        public readonly TipologiaCartelle TipologiaCartella;
        public readonly NomiDatoErrore NomeDatoErrore;
        public readonly string WorksheetName;
        public readonly string PercorsoFile;
        public readonly int? RigaCella;
        public readonly int? ColonnaCella;
        public readonly string NomeColonnaCella;
        public readonly string Dato;

        private string _messaggioPerUtente;

        public string MessaggioPerUtente
        {
            get
            {
                if (string.IsNullOrEmpty(_messaggioPerUtente))
                {
                    _messaggioPerUtente = $"{TipologiaErrore.GetEnumDescription()}, cartella: {TipologiaCartella.GetEnumDescription()}";

                    if (!string.IsNullOrEmpty(WorksheetName))
                    {
                        _messaggioPerUtente += $", foglio: \"{WorksheetName}\"";
                    }

                    if (ColonnaCella.HasValue && RigaCella.HasValue)
                    {
                        _messaggioPerUtente += $", cella: {NomeColonnaCella}{RigaCella}";
                    }
                    else
                    {
                        if (ColonnaCella.HasValue)
                        {
                            _messaggioPerUtente += $", colonna: {NomeColonnaCella}";
                        }

                        if (RigaCella.HasValue)
                        {
                            _messaggioPerUtente += $", riga: {RigaCella}";
                        }
                    }

                    if (NomeDatoErrore != NomiDatoErrore.None)
                    {
                        _messaggioPerUtente += $", dato: {NomeDatoErrore.GetEnumDescription()}";
                    }

                    if (!string.IsNullOrEmpty(Dato))
                    {
                        _messaggioPerUtente += $", valore: \"{Dato}\"";
                    }
                }

                return _messaggioPerUtente;
            }
        }

        public ManagedException(
                TipologiaErrori tipologiaErrore,
                TipologiaCartelle tipologiaCartella,
                string messaggioPerUtente = null,
                string worksheetName = null,
                string percorsoFile = null,
                int? rigaCella = null,
                int? colonnaCella = null,
                string dato = null,
                NomiDatoErrore nomeDatoErrore = NomiDatoErrore.None
                ) : base(messaggioPerUtente)
        {
            _messaggioPerUtente = messaggioPerUtente;

            TipologiaErrore = tipologiaErrore;
            TipologiaCartella = tipologiaCartella;
            WorksheetName = worksheetName;
            PercorsoFile = percorsoFile;
            RigaCella = rigaCella;
            ColonnaCella = colonnaCella;
            NomeColonnaCella = colonnaCella.HasValue 
                            ? ((ColumnIDS)colonnaCella.Value).ToString()
                            : null;
            Dato = dato;
            NomeDatoErrore = nomeDatoErrore;
        }
    }
}