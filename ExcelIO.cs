/* 
 * Copyright (c) 2013, Gareth Higgins
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
    * Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.
    * Neither the name of the <organization> nor the
      names of its contributors may be used to endorse or promote products
      derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * 
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace WineExcel {
    class ExcelIO : IDisposable {
        private const string StartingCellProducers = "A6";
        private const string SearchCellProducers = "B6";
        private const int StartingRowOffsetProducers = 5;
        private const string StartingCellTemplate = "A2";
        private const int StaticDataFieldCount = 10;
        private const int StaticDataFieldCountExtended = 0;
        private const int ExtendedDataReadEnd = StaticDataFieldCount + StaticDataFieldCountExtended;

        private Application _excel;
        private Workbooks _workbooks;
        private Workbook _wTemplate;
        private Worksheet[] _worksheet;

        private Range _rObj;
        private bool _isDisposed;

        private readonly string _executingPath;
        private const string TemplatePath = "template.xlsx";
        private readonly string _loadPathTemplate;
        private const string ProducersPath = "producers.xls";
        private readonly string _loadPathProducers;
        private const string OutputFileName = "SLAVE";
        private readonly string _outPath;

        public ExcelIO()
        {
            // ReSharper disable AssignNullToNotNullAttribute
            _executingPath = new Uri(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase)).LocalPath;
            _loadPathTemplate = Path.Combine(_executingPath, TemplatePath);
            _loadPathProducers = Path.Combine(_executingPath, ProducersPath);
            _outPath = Path.Combine(_executingPath, OutputFileName);
            // ReSharper restore AssignNullToNotNullAttribute
#if DEBUG
            Console.WriteLine("Preparing excel application");
            _excel = new Application { Visible = true, ScreenUpdating = true, DisplayAlerts = true };
#else
            _excel = new Application {Visible = false, ScreenUpdating = false, DisplayAlerts = false};
#endif
            _isDisposed = false;

        }

        public void OpenProducers() {
#if DEBUG
            Console.WriteLine("Opening up excel producer file");
#endif
            if (_excel == null || _isDisposed)
                throw new ObjectDisposedException("excel", "Either the excel app has been closed, or it was never opened");
            if (!File.Exists(_loadPathProducers)) 
                throw new FileNotFoundException(string.Format("Unable to find producer file at path: {0}", _loadPathProducers));
            
            _workbooks = _excel.Workbooks;
            _wTemplate = _workbooks.Open(_loadPathProducers, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            _worksheet = new Worksheet[] { _wTemplate.Worksheets[1] };
        }

        public Producers ReadProducers() {
            Producers pd = new Producers();
            Producer dummy = null;

#if DEBUG
            Console.WriteLine("Reading producers");
#endif

            _rObj = _worksheet[0].Range[SearchCellProducers, Missing.Value];
            _rObj = _rObj.End[XlDirection.xlDown];
            string downAddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _worksheet[0].Range[StartingCellProducers, "A" + downAddress.Substring(1)];
            object[,] values = (object[,])_rObj.Value2;
            for (int i = 1; i <= values.GetLength(0); i++) {
                if ((string)values[i,1] != null) {
                    dummy = new Producer();
                    dummy.ParseFromString((string)values[i,1]);
                    pd.AddProducer(dummy);
                }
                Debug.Assert(dummy != null, "dummy != null");
                pd.AddProduct(dummy.ID, _worksheet[0].Cells[i + StartingRowOffsetProducers, 2].Value2);
            }
            
#if DEBUG
            Console.WriteLine("Successfully added {0} producer{1} with {2} product{3}",pd.Count, pd.Count != 0 ? "s" : "", pd.ProductCount, pd.ProductCount != 0 ? "s" : "");
#endif
            Marshal.ReleaseComObject(_rObj);
            _rObj = null;
            return pd;
        }

        public void OpenTemplate()
        {
#if DEBUG
            Console.WriteLine("Opening up excel template file");
#endif
            if (_excel == null || _isDisposed)
                throw new ObjectDisposedException("excel", "Either the excel app has been closed, or it was never opened");
            if (!File.Exists(_loadPathTemplate))
                throw new FileNotFoundException(string.Format("Unable to find template file at path: {0}", _loadPathTemplate));
            _workbooks = _excel.Workbooks;
            _wTemplate = _workbooks.Open(_loadPathTemplate, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            _worksheet = new Worksheet[] { _wTemplate.Worksheets[1], _wTemplate.Worksheets[2], _wTemplate.Worksheets[3] };
        }

        public void SetupDocument(Dictionary<int, Product> products)
        {
            if (_isDisposed)
                throw new ObjectDisposedException("excel", "Either the excel app has been closed, or it was never opened");
#if DEBUG
            Console.WriteLine("Writing Worksheet Names");
#endif
            _worksheet[0].Name = "Template";
            _worksheet[1].Name = "Listing";
            _worksheet[2].Name = "Delisted";

#if DEBUG
            Console.WriteLine("Copying Static Headers");
#endif
            _rObj = _worksheet[0].Range["A1", GetExcelColumnName(StaticDataFieldCount) + 1];
            _rObj.Copy(_worksheet[1].Range["A1", GetExcelColumnName(StaticDataFieldCount) + 1]);
            _rObj.Copy(_worksheet[2].Range["A1", GetExcelColumnName(StaticDataFieldCount) + 1]);

            if (StaticDataFieldCountExtended > 0)
            {
                _rObj =
                    _worksheet[0].Range[
                        GetExcelColumnName(StaticDataFieldCount + 1) + 1, GetExcelColumnName(ExtendedDataReadEnd) + 1];
                _rObj.Copy(
                    _worksheet[1].Range[
                        GetExcelColumnName(StaticDataFieldCount + products.Count + 2) + 1,
                        GetExcelColumnName(ExtendedDataReadEnd + products.Count + 1) + 1]);
            }
            int q = 1;

#if DEBUG
            Console.WriteLine("Writing Out Product Headers");
#endif

            foreach (var product in products) {
                if (product.Value.Name != null) {
                    if (product.Value.ID < 0) {
                        _worksheet[1].Cells[1, StaticDataFieldCount + q] = product.Value.Name;
                    } else {
                        _worksheet[1].Cells[1, StaticDataFieldCount + q] = string.Format("{0}\n({1})", product.Value.Name,
                                                                           product.Key);
                    }

                } else {
                    _worksheet[1].Cells[1, StaticDataFieldCount + q] = product.Key;
                }

                q++;
            }
            _worksheet[1].Cells[1, + products.Count + StaticDataFieldCount + 1] = "Date Updated";
            _worksheet[2].Cells[1, StaticDataFieldCount + 1] = "Date Updated";
        }

        public void WriteProductDataToStores(Dictionary<int, Dictionary<int,ProductEntry>> sPp, Dictionary<int, Product> products , Producers pList)
        {
            if (_isDisposed)
                throw new ObjectDisposedException("excel", "Either the excel app has been closed, or it was never opened");
#if DEBUG
            Console.WriteLine("Getting Store Data From Template");
#endif  
            _rObj = _worksheet[0].Range[StartingCellTemplate, Missing.Value];
            _rObj = _rObj.End[XlDirection.xlDown];
            string downAddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _worksheet[0].Range[StartingCellTemplate, downAddress];
            object[,] values = (object[,])_rObj.Value2;
#if DEBUG
            Console.WriteLine("Writing ProductInformation to Stores");
#endif      
            int q;
            int lineList = 1;
            int lineDeList = 1;
            for (int i = 1; i <= values.GetLength(0); i++) {
                int key = Convert.ToInt32(values[i, 1]);
                if (!sPp.ContainsKey(key)) continue;
                bool listed = sPp[key].Any(entry => entry.Value.ListingState == ListingStatus.Listed || entry.Value.ListingState == ListingStatus.Forced);
                q = 1;
                if (listed) {
                    foreach (var product in products) {
                        if (sPp[key].ContainsKey(product.Key)) {
                                _worksheet[1].Cells[lineList + 1, StaticDataFieldCount + q] = sPp[key][product.Key].InventoryChange;
                            _worksheet[1].Cells[lineList + 1, StaticDataFieldCount + q].Interior.Color =
                                pList.GetProducerColourCodeFromProduct(product.Key);
                        }
                        q++;
                    }
                }
                _rObj = _worksheet[0].Range["A" + (i + 1), GetExcelColumnName(StaticDataFieldCount) + (i + 1)];
                if (listed) {
                    _rObj.Copy(_worksheet[1].Range["A" + (lineList + 1), GetExcelColumnName(StaticDataFieldCount) + (lineList + 1)]);
                    _worksheet[1].Cells[lineList + 1, StaticDataFieldCount + products.Count + 1] = DateTime.Now.ToShortDateString();
                    if (StaticDataFieldCountExtended > 0)
                    {
                        _rObj =
                            _worksheet[0].Range[
                                GetExcelColumnName(StaticDataFieldCount + 1) + (lineList + 1),
                                GetExcelColumnName(ExtendedDataReadEnd) + (lineList + 1)];

                        _rObj.Copy(_worksheet[1].Range[
                            GetExcelColumnName(StaticDataFieldCount + products.Count + 2) + (lineList + 1),
                            GetExcelColumnName(ExtendedDataReadEnd + products.Count + 1) + (lineList + 1)]);
                    }

                    lineList++;
                } else {
                    _rObj.Copy(_worksheet[2].Range["A" + (lineDeList + 1), GetExcelColumnName(StaticDataFieldCount) + (lineDeList + 1)]);

                    _worksheet[2].Cells[lineDeList + 1, StaticDataFieldCount + 1] = DateTime.Now.ToShortDateString();
                    lineDeList++;
                }
            }
            _rObj = _worksheet[1].Cells[2, StaticDataFieldCount + products.Count + 1];
            string upaddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _rObj.End[XlDirection.xlDown];
            downAddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _worksheet[1].Range[upaddress, downAddress];
            _rObj.EntireColumn.NumberFormat = "m/d/yyyy";
            _rObj = _worksheet[2].Cells[2, StaticDataFieldCount + 1];
            upaddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _rObj.End[XlDirection.xlDown];
            downAddress = _rObj.Address[false, false, XlReferenceStyle.xlA1, Missing.Value, Missing.Value];
            _rObj = _worksheet[2].Range[upaddress, downAddress];
            _rObj.EntireColumn.NumberFormat = "m/d/yyyy";
        }

            

        public void CleanupAndSaveCopy()
        {
            if (_isDisposed)
                throw new ObjectDisposedException("excel", "Either the excel app has been closed, or it was never opened");
#if DEBUG
            Console.WriteLine("Removing Template and Saving");
#endif 
            ((Worksheet)_wTemplate.Worksheets[1]).Delete();



            _wTemplate.SaveAs(_outPath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                            false, false, XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(_rObj);
            _rObj = null;
        }

        public void Dispose() {
#if DEBUG
            Console.WriteLine("Cleaning up");
#endif
            _isDisposed = true;
            if (_worksheet != null) {
                for (int i = 0; i < _worksheet.Length; i++) {
                    if (_worksheet[i] != null) {
                        Marshal.ReleaseComObject(_worksheet[i]);
                    }
                    _worksheet[i] = null;
                }
            }

            _worksheet = null;
            
            if (_wTemplate != null)
            {
                _wTemplate.Close(false, Missing.Value, Missing.Value);
                Marshal.ReleaseComObject(_wTemplate);
            }
            _wTemplate = null;
            if (_workbooks != null)
            {
                _workbooks.Close();
                Marshal.ReleaseComObject(_workbooks);
            }
            _workbooks = null;
            if (_excel != null)
            {
                _excel.Quit();
                Marshal.ReleaseComObject(_excel);

            }
            _excel = null;
        }

        private static string GetExcelColumnName(int columnNumber) {
            int dividend = columnNumber;
            string columnName = String.Empty;

            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString(CultureInfo.InvariantCulture) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
