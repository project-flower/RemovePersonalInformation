using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace RemovePersonalInformation
{
    public static class MainEngine
    {
        #region Private Fields

        private static Documents documents = null;
        private static Excel.Application excelApplication = null;
        private static PowerPoint.Application powerPointApplication = null;
        private static Presentations presentations = null;
        private static Word.Application wordApplication = null;
        private static Workbooks workbooks = null;

        #endregion

        #region Public Methods

        public static void RemovePersonalInformation(string[] fileNames, out string[] completedFiles, out string[] faultFiles)
        {
            completedFiles = null;
            faultFiles = null;

            try
            {
                var completedFilesList = new List<string>();
                var faultFilesList = new List<string>();

                foreach (string fileName in fileNames)
                {
                    try
                    {
                        switch (Path.GetExtension(fileName))
                        {
                            case ".doc":
                            case ".docx":
                                removeFromWord(fileName);
                                break;

                            case ".ppt":
                            case ".pptx":
                                removeFromPowerPoint(fileName);
                                break;

                            case ".xls":
                            case ".xlsm":
                            case ".xlsx":
                                removeFromExcel(fileName);
                                break;

                            default:
                                faultFilesList.Add(fileName);
                                continue;
                        }

                        completedFilesList.Add(fileName);
                    }
                    catch
                    {
                        faultFilesList.Add(fileName);
                        continue;
                    }
                }

                completedFiles = completedFilesList.ToArray();
                faultFiles = faultFilesList.ToArray();
            }
            finally
            {
                if (workbooks != null)
                {
                    try
                    {
                        // "通常使うプログラムとして設定されていません。"が表示されている場合、例外がスローされる。
                        workbooks.Close();
                    }
                    catch
                    {
                    }

                    releaseComObject(ref workbooks);
                }

                if (documents != null)
                {
                    try
                    {
                        // "通常使うプログラムとして設定されていません。"が表示されている場合、例外がスローされる。
                        documents.Close();
                    }
                    catch
                    {
                    }

                    releaseComObject(ref documents);
                }

                if (presentations != null)
                {
                    releaseComObject(ref presentations);
                }

                GC.Collect();

                if (excelApplication != null)
                {
                    try
                    {
                        // "サブスクリプションの有効期限が切れています"が表示されている場合、例外がスローされる。
                        excelApplication.Quit();
                    }
                    catch
                    {
                    }

                    releaseComObject(ref excelApplication);
                }

                if (wordApplication != null)
                {
                    try
                    {
                        // "サブスクリプションの有効期限が切れています"が表示されている場合、例外がスローされる。
                        wordApplication.Quit();
                    }
                    catch
                    {
                    }

                    releaseComObject(ref wordApplication);
                }

                if (powerPointApplication != null)
                {
                    try
                    {
                        // "サブスクリプションの有効期限が切れています"が表示されている場合、例外がスローされる。
                        powerPointApplication.Quit();
                    }
                    catch
                    {
                    }

                    releaseComObject(ref powerPointApplication);
                }

                GC.Collect();
            }
        }

        #endregion

        #region Private Methods

        private static void releaseComObject<T>(ref T o)
        {
            if (o == null)
            {
                return;
            }

            Marshal.FinalReleaseComObject(o);
            o = default;
        }

        private static void removeFromExcel(string fileName)
        {
            if (excelApplication == null)
            {
                try
                {
                    excelApplication = new Excel.Application();
                    excelApplication.DisplayAlerts = false;
                }
                catch
                {
                    throw;
                }
            }

            if (workbooks == null)
            {
                try
                {
                    workbooks = excelApplication.Workbooks;
                }
                catch
                {
                    throw;
                }
            }

            Workbook workbook = null;

            try
            {
                workbook = workbooks.Open(fileName);
                workbook.RemoveDocumentInformation(XlRemoveDocInfoType.xlRDIRemovePersonalInformation);
                workbook.Save();
            }
            catch
            {
                throw;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                    workbook = null;
                }

                releaseComObject(ref workbook);
            }
        }

        private static void removeFromPowerPoint(string fileName)
        {
            if (powerPointApplication == null)
            {
                try
                {
                    powerPointApplication = new PowerPoint.Application();
                    powerPointApplication.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch
                {
                    throw;
                }
            }

            if (presentations == null)
            {
                try
                {
                    presentations = powerPointApplication.Presentations;
                }
                catch
                {
                    throw;
                }
            }

            Presentation presentation = null;

            try
            {
                presentation = presentations.Open(fileName);
                presentation.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIRemovePersonalInformation);
                presentation.Save();
            }
            catch
            {
                throw;
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                }

                releaseComObject(ref presentation);
            }
        }

        private static void removeFromWord(string fileName)
        {
            if (wordApplication == null)
            {
                try
                {
                    wordApplication = new Word.Application();
                    wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                    throw;
                }
            }

            if (documents == null)
            {
                try
                {
                    documents = wordApplication.Documents;
                }
                catch
                {
                    throw;
                }
            }

            Document document = null;

            try
            {
                document = documents.Open(fileName);
                document.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIRemovePersonalInformation);
                document.Save();
            }
            catch
            {
                throw;
            }
            finally
            {
                if (document != null)
                {
                    document.Close();
                }

                releaseComObject(ref document);
            }
        }

        #endregion
    }
}
