using EPPlusExtensions;
using System;

namespace ReportRefresher.Entities
{
    public abstract class InfoFileBase
    {
        public InfoFileBase(EPPlusHelper epPlusHelper, string filePath)
        {
            if (epPlusHelper == null)
                throw new ArgumentNullException(nameof(epPlusHelper));
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            EPPlusHelper = epPlusHelper;
            FilePath = filePath;
        }

        public InfoFileBase()
        {
        }

        public string FilePath
        {
            get; private set;
        }

        public EPPlusHelper EPPlusHelper
        {
            get; private set;
        }
    }
}