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
    internal class RawReader
    {
        private const char CommentLine = ';';
        private Dictionary<int, Product> _products;

        private Dictionary<int, Dictionary<int, ProductEntry>> _storeProductPair;

        public void ReadProducts(string filename)
        {
#if DEBUG
            Console.WriteLine("Reading current product data");
#endif
            string[] temp = File.ReadAllLines(filename);
#if DEBUG
            Console.WriteLine("Processing {0} lines", temp.Length);
#endif
            _products = new Dictionary<int, Product>(temp.Length);
            foreach (string t in temp)
            {
                if (t.Length > 0 && t[0] == CommentLine) continue;

                Product dummy = new Product();
                dummy.SetProductIDFromRaw(t);
                if (_products.ContainsKey(dummy.ID)) {
                    throw new ArgumentException(string.Format("An item with the same product ID already exist. This: {0}({1}), Found: {2}({3})", dummy.ID, dummy.Name, _products[dummy.ID].ID, _products[dummy.ID].Name));
                }
                _products.Add(dummy.ID, dummy);
            }
#if DEBUG
            Console.WriteLine("{0} products successfully added", _products.Count);
#endif
        }

        public void ReadFile(string filename)
        {
#if DEBUG
            Console.WriteLine("Reading LCBO inventory update");
#endif
            if (_storeProductPair == null)
            {
                _storeProductPair = new Dictionary<int, Dictionary<int, ProductEntry>>();
            }
            if (_products == null || _products.Count < 1)
            {
                throw new InvalidDataException(
                    "No products were specified to be searched for, did you forget to call ReadProducts?");
            }
            using (TextReader textReader = new StreamReader(filename, false))
            {
                string line;
                while ((line = textReader.ReadLine()) != null)
                {
                    int pID;
                    if (!int.TryParse(line.Substring(8, 7), out pID))
                    {
                        throw new InvalidDataException(string.Format("Unable to parse Product ID '{0}'",
                                                                     line.Substring(8, 7)));
                    }

                    //Console.WriteLine(line.Substring(0, 8));
                    //Console.WriteLine(line.Substring(8, 7));

                    if (!_products.ContainsKey(pID)) continue;

                    //if (!_products.Contains()) continue;

                    ProductEntry pEnt = new ProductEntry();
                    pEnt.SetInfoFromRaw(line);

                    if (!_storeProductPair.ContainsKey(pEnt.StoreID))
                    {
                        _storeProductPair.Add(pEnt.StoreID, new Dictionary<int, ProductEntry>());
                    }
                    _storeProductPair[pEnt.StoreID].Add(pEnt.ProductID, pEnt);
                }
            }
#if DEBUG
            int i = 0;
            foreach (var spp in _storeProductPair)
            {
                if (spp.Value.Count > i)
                {
                    i = spp.Value.Count;
                }
            }
            Console.WriteLine("Found {0} products in {1} store{2}", i, _storeProductPair.Count, _storeProductPair.Count != 1 ? "s": "");
#endif
        }

        public void ResetStoreProductPairs()
        {
            _storeProductPair.Clear();
        }

        public Dictionary<int, Dictionary<int, ProductEntry>> GetStoreProductEntries()
        {
            return new Dictionary<int, Dictionary<int, ProductEntry>>(_storeProductPair);
        }

        public Dictionary<int, Product> GetProducts()
        {
            return new Dictionary<int, Product>(_products);
        }
    }
}
