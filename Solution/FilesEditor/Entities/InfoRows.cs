using System;

namespace FilesEditor.Entities
{
    public class InfoRows
    {
        IntSetOnce<int> _iniziali;
        IntSetOnce<int> _preservate;
        IntSetOnce<int> _riutilizzate;
        IntSetOnce<int> _eliminate;
        IntSetOnce<int> _aggiunte;
        IntSetOnce<int> _finali;

        public int Iniziali
        {
            get { return _iniziali.Value; }
            set { _iniziali.Value = value; }
        }

        public int Preservate
        {
            get { return _preservate.Value; }
            set { _preservate.Value = value; }
        }

        public int Riutilizzate
        {
            get { return _riutilizzate.Value; }
            set { _riutilizzate.Value = value; }
        }

        public int Eliminate
        {
            get { return _eliminate.Value; }
            set { _eliminate.Value = value; }
        }

        public int Aggiunte
        {
            get { return _aggiunte.Value; }
            set { _aggiunte.Value = value; }
        }

        public int Finali
        {
            get { return _finali.Value; }
            set { _finali.Value = value; }
        }


        //todo: eliminabile...
        public void VerificaCoerenzaValori()
        {
            // mi assicuro che siano stati tutti valorizzati
            var sommaDiTutteLeVariabili =
                Iniziali +
                Preservate +
                Riutilizzate +
                Eliminate +
                Aggiunte +
                Finali;

            if (Finali != Iniziali - Eliminate + Aggiunte)
            {
                throw new Exception($"Errore durante la verifica del numero di righe finali. Finali: {Finali} - Iniziali: {Iniziali} - Eliminate: {Eliminate} - Aggiunte {Aggiunte}");
            }
        }
    }
}
