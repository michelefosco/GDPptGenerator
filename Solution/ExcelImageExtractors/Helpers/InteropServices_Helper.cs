using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;


namespace ExcelImageExtractors.Helpers
{
    internal class InteropServices_Helper
    {
        private const uint RPC_E_CALL_REJECTED = 0x80010001;

        static internal void PrioritizeExcelProcess()
        {
            var excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL").FirstOrDefault();
            if (excelProcess != null)
            {
                excelProcess.PriorityClass = System.Diagnostics.ProcessPriorityClass.High;
            }
        }

        static internal void RetryComCall(Action action, string errorMessageToIgnore = null)
        {
            const int ADDED_WAITING_TIME_MS = 100;
            const int MIMINUM_WAITING_TIME_MS = 300;

            var waitingTimeMs = MIMINUM_WAITING_TIME_MS;
            var availableAttempts = 20;

            while (availableAttempts-- > 0)
            {
                try
                {
                    action();
                    return; //or break;
                }
                catch (COMException ex)// when ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                {
                    if (availableAttempts > 1)
                    {
                        // 1) condizione da ignorare
                        if ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                        {
                            // Excel is  busy, attendo per qualche millisecondo
                            Thread.Sleep(waitingTimeMs += ADDED_WAITING_TIME_MS);
                            continue;
                        }

                        // 2° condizione da ignorare
                        if (!string.IsNullOrEmpty(errorMessageToIgnore) && ex.Message.Equals(errorMessageToIgnore, StringComparison.Ordinal))
                        {
                            // Per gli errori da ignorare applico il tempo minimo di delay
                            Thread.Sleep(MIMINUM_WAITING_TIME_MS);
                            continue;
                        }
                    }

                    throw;
                }
            }

            throw new TimeoutException("Excel did not become ready in time.");
        }
    }
}
