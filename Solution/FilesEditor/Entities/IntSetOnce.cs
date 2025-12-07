using System;

namespace FilesEditor.Entities
{
    internal struct IntSetOnce<T> where T : struct
    {
        Nullable<T> _value;
        public T Value
        {
            get
            {
                if (!_value.HasValue)
                { throw new Exception("Il valore non è stato ancora impostato."); }
                return _value.Value;
            }
            set
            {
                if (_value.HasValue)
                { throw new Exception("Questa variabile è di tipo SetOnce. Il valore è già stato impostato. Quindi non può essere modificato"); }
                _value = value;
            }
        }
    }
}
