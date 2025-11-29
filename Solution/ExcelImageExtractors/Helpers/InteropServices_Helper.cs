using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;


namespace ExcelImageExtractors.Helpers
{
    internal class InteropServices_Helper
    {
        const uint RPC_E_CALL_REJECTED = 0x80010001;

        static internal void PrioritizeExcelProcess()
        {
            var excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL").FirstOrDefault();
            if (excelProcess != null)
            { 
                excelProcess.PriorityClass = System.Diagnostics.ProcessPriorityClass.High; 
            }
        }

        static internal  void WaitForExcelReady(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            int maxWait = 200;  // max 20 seconds
            int waitingTimeMs = 100;
            while (maxWait-- > 0)
            {
                try
                {
                    // Attempt a harmless call to check if Excel is responsive
                    var test = excelApp.Hwnd;
                    return; // no exception → Excel is ready
                }
                catch (COMException ex)
                {
                    // Excel is temporarily busy (edit mode, dialog, recalculating)
                    if ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                    {
                        Thread.Sleep(waitingTimeMs += 100);
                        continue;
                    }

                    // If it's not RPC_E_CALL_REJECTED → rethrow
                    throw;
                }
            }

            throw new TimeoutException("Excel did not become ready in time.");
        }

        static internal void RetryComCall(System.Action action)
        {

            int retries = 15;

            int waitingTimeMs = 100;

            while (true)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                {
                    if (--retries == 0)
                    { throw; }

                    Thread.Sleep(waitingTimeMs += 100);
                }
            }
        }
    }
}
