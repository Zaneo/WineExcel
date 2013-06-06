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

using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System;


namespace WineExcel {

    [DebuggerDisplay("Product = {ID}({Name})")]
    class Product 
    {
        static readonly char[] CharSep = new[]{':'};
        public int ID { get; private set; }
        public string Name { get; private set; }
        public DateTime YearProduced { get; private set; }

        private static int _uniqueFakeProductID = 0;

        public Product()
        {
        }

        public static int GetFakeProductID()
        {
            if (_uniqueFakeProductID == int.MinValue)
                throw new KeyNotFoundException("Ran out of fake product IDs to assign");
            return --_uniqueFakeProductID;
        }

        public Product(int pID)
        {
            ID = pID;
        }

        public void SetProductIDFromRaw(string raw)
        {
            string[] strings = raw.Split(CharSep);
            int pid;
            if(!int.TryParse(strings[0], out pid))
            {
                throw new InvalidDataException(string.Format("Unable to parse '{0}' expected integer", raw));
            }
            ID = pid;
            if (ID < 0)
            {
                ID = GetFakeProductID();
            }
            if (strings.Length > 1)
            {
                //string.Replace("\\n", "\n");
                Name = strings[1].Trim();
            }
            if (strings.Length > 2) {
                
            }
        }

        public override string ToString()
        {
            return ID.ToString(CultureInfo.InvariantCulture);
        }
    }
}
