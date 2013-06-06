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
using System.IO;

namespace WineExcel {
    class Program {
        static readonly RawReader RawReader = new RawReader();
        private static ExcelIO _eIo;
        static Dictionary<int, Dictionary<int,ProductEntry>> _storeProducts;
        private static Dictionary<int, Product> _products;
        private static Producers _producers;
        static void Main(string[] args) {

            try
            {
                Console.Title = "WineExcel";
#if DEBUG
                Console.ForegroundColor = ConsoleColor.DarkMagenta;
#endif
                RawReader.ReadProducts("productMapping.txt");
                RawReader.ReadFile("data.dat");
                _storeProducts = RawReader.GetStoreProductEntries();
                _products = RawReader.GetProducts();
                RawReader.ResetStoreProductPairs();

                using (_eIo = new ExcelIO()) {
                    _eIo.OpenProducers();
                    _producers = _eIo.ReadProducers();
                }
                using (_eIo = new ExcelIO()) {
                    _eIo.OpenTemplate();
                    _eIo.SetupDocument(_products);
                    _eIo.WriteProductDataToStores(_storeProducts,_products, _producers);
                    _eIo.CleanupAndSaveCopy();
                }

            }
            catch (Exception e) {
                Console.WriteLine("An error occurred somewhere, a log will be written to 'ERROR" + DateTime.UtcNow.ToFileTimeUtc() + ".txt'. Please infom the developer");
                try {
                    File.WriteAllText("ERROR-" + DateTime.UtcNow.ToFileTimeUtc() + ".txt", e.Message+"\n" + e.StackTrace);
                }
                catch (Exception e2) {
                    Console.WriteLine("Writing the error log failed: " + e2.Message + "\nOriginal: " + e.Message + "\n" + e.StackTrace);
#if DEBUG
                    throw e2;
#endif
                }
#if DEBUG
                throw e;
#endif
            }
        }
    }
}
