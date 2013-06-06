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
using System.Globalization;
using System.IO;


namespace WineExcel {
    enum ListingStatus
    {
        Listed,
        Delisted,
        Forced,
    }

    internal class ProductEntry
    {
        private const string DatePattern = "yyyyMMdd";
        public DateTime EntryTime { get; private set; }
        public int ProductID { get; private set; }
        public int StoreID { get; private set; }
        public int InventoryChange { get; private set; }
        public ListingStatus ListingState { get; private set; }


        public ProductEntry()
        {
        }

        public ProductEntry(DateTime date, int pId, int sId, int invent, ListingStatus lS)
        {
            EntryTime = date;
            ProductID = pId;
            StoreID = sId;
            InventoryChange = invent;
            ListingState = lS;
        }

        public void SetInfoFromRaw(string rawData)
        {
            DateTime tempTime;
            int tempInt;
            if (!DateTime.TryParseExact(rawData.Substring(0, 8), DatePattern, null, DateTimeStyles.None, out tempTime))
            {
                throw new InvalidDataException(string.Format("Date '{0}' not in parseable format '{1}'",
                                                             rawData.Substring(0, 8), DatePattern));
            }
            EntryTime = tempTime;
            if (!int.TryParse(rawData.Substring(8, 7), out tempInt))
            {
                throw new InvalidDataException(string.Format("ID '{0} not in paresable format",
                                                             rawData.Substring(8, 8)));
            }
            ProductID = tempInt;
            if (!int.TryParse(rawData.Substring(15, 4), out tempInt))
            {
                throw new InvalidDataException(string.Format("StoreID '{0} not in paresable format",
                                                             rawData.Substring(16, 3)));
            }
            StoreID = tempInt;
            switch (rawData[19])
            {
                case 'D':
                    ListingState = ListingStatus.Delisted;
                    break;
                case 'L':
                    ListingState = ListingStatus.Listed;
                    break;
                case 'F':
                    ListingState = ListingStatus.Forced;
                    break;
                default:
                    throw new InvalidDataException(string.Format("Listing Status '{0} not in paresable format",
                                                                 rawData[19]));
            }
            if (!int.TryParse(rawData.Substring(20), out tempInt))
            {
                throw new InvalidDataException(string.Format("Inventory '{0} not in paresable format",
                                                             rawData.Substring(19)));
            }
            InventoryChange = tempInt;
        }
    }
}
