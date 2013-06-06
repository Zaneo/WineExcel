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
using System.Drawing;

namespace WineExcel
{
    internal class Producers {

        public int Count { get { return _producers.Count; } }
        public int ProductCount { get { return _claimedProducts.Count; } }

        private readonly Dictionary<int, Producer> _producers;
        private readonly Dictionary<int, int> _claimedProducts;
        private readonly Stack<Color> _availableColours;
 
        public Producers() {
            _producers = new Dictionary<int, Producer>();
            _claimedProducts = new Dictionary<int, int>();
            _availableColours = new Stack<Color>(50);

            #region Colours
                _availableColours.Push(Color.DarkTurquoise);
                _availableColours.Push(Color.DarkSeaGreen);
                _availableColours.Push(Color.DarkViolet);
                _availableColours.Push(Color.DarkSalmon);
                _availableColours.Push(Color.DarkKhaki);
                _availableColours.Push(Color.DarkGray);
                _availableColours.Push(Color.Cyan);
                _availableColours.Push(Color.DeepPink);
                _availableColours.Push(Color.GreenYellow);
                _availableColours.Push(Color.MediumPurple);
                _availableColours.Push(Color.RosyBrown);
                _availableColours.Push(Color.Tomato);
                _availableColours.Push(Color.Aquamarine);
                _availableColours.Push(Color.ForestGreen);
                _availableColours.Push(Color.Blue);
                _availableColours.Push(Color.DeepSkyBlue);
                _availableColours.Push(Color.Khaki);
                _availableColours.Push(Color.DarkOrange);
                _availableColours.Push(Color.Gold);
                _availableColours.Push(Color.Green);
                _availableColours.Push(Color.LightCoral);
                _availableColours.Push(Color.LightGreen);
                _availableColours.Push(Color.LightSeaGreen);
                _availableColours.Push(Color.MediumOrchid);
                _availableColours.Push(Color.OrangeRed);
                _availableColours.Push(Color.MediumSpringGreen);
                _availableColours.Push(Color.PaleVioletRed);

            #endregion
        }

        public void AddProducer(Producer pd) {
            if (_producers.ContainsKey(pd.ID)) 
                throw new ArgumentException(string.Format("ProducerID ({0}) already belongs to another producer: {1}({2})", pd.ID, _producers[pd.ID].ID, _producers[pd.ID].Name));
            
            if (pd.DisplayColour != Color.Empty) {
                throw new AccessViolationException("Producer likely belongs to another set of producers and forgot to return it's display colour, call Producers.RemoveProducer() first.");
            }
            if (_availableColours.Count == 0) 
                throw new InvalidOperationException("There are no more colours to give out, either remove some producers, or ask the dev to add more colours.");

            pd.SetColour(_availableColours.Pop());
            _producers.Add(pd.ID, pd);
        }

        public void AddProduct(int producer, int productID)
        {
            if (_claimedProducts.ContainsKey(productID))
            {
                int prID = _claimedProducts[productID];
                throw new ArgumentException(string.Format("Product ({0}) already belongs to another producer: {1}({2})", productID, prID, _producers[prID].Name));
            }
            _producers[producer].AddProduct(productID);
            _claimedProducts.Add(productID, producer);
        }

        public void AddProduct(int producer, string raw) {
            int idEnd = raw.IndexOf(' ');
            int pid;
            if (idEnd < 0)
            {
                throw new ArgumentException(string.Format("Malformed string: {0}", raw));
            }
            if (!int.TryParse(raw.Substring(0, idEnd), out pid))
            {
                throw new ArgumentException(string.Format("Unable to parse product ID: {0}", raw.Substring(0, idEnd)));
            }
            AddProduct(producer, pid);
        }

        public Producer GetProducerForProduct(int id)
        {
            if (!_claimedProducts.ContainsKey(id))
            {
                throw new ArgumentException(string.Format("No product was found with that ID: {0}", id));
            }
            return (Producer) _producers[_claimedProducts[id]].Clone();
        }

        public Color GetProducerColourCodeFromProduct(int productID) {
            int pid;
            if (!_claimedProducts.TryGetValue(productID,out pid))
                throw new ArgumentException(string.Format("Unable to find product under any producer: {0}", productID));
            return GetProducerColourCodeFromID(pid);
        }

        public Color GetProducerColourCodeFromID(int producerID) {
            Producer pd;
            if(!_producers.TryGetValue(producerID, out pd))
                throw new ArgumentException(string.Format("Unable to find producer: {0}", producerID));
            return pd.DisplayColour;
        }

        public void RemoveProduct(int producer, int productID)
        {
            int prID;
            if (!_claimedProducts.TryGetValue(productID, out prID)) {
                throw new ArgumentException(
                    string.Format("Could not find specified product ({0}) under this producer or any other", productID));
            }
            if (prID != productID) {
                throw new ArgumentException(
                    string.Format(
                        "Could not find specified product ({0}) under this producer, but it belongs to another producer: {1}({2})",
                        productID, prID, _producers[prID].Name));
            }
            _producers[prID].RemoveProduct(productID);
        }

        public void RemoveProducer(int producerID)
        {
            Producer pd;
            if (!_producers.TryGetValue(producerID, out pd))
            {
                throw new ArgumentException(string.Format("Unable to find a producer with the specified id: {0}", producerID));
            }
            foreach (var pro in pd.GetProductIDs())
            {
                _claimedProducts.Remove(pro);
            }
            _availableColours.Push(pd.DisplayColour);
            pd.SetColour(Color.Empty);
            _producers.Remove(producerID);
        }
    }

    public class Producer : ICloneable {
        public int ID { get; private set; }
        public string Name { get; private set; }

        private readonly List<int> _products;

        public Color DisplayColour { get; private set; }

        public Producer() : this(default(int), null, Color.Empty){}

        public Producer(int id, string name) : this(id, name, Color.Empty) { }

        public Producer(int id, string name, Color displayColour)
        {
            ID = id;
            Name = name;
            _products = new List<int>();
            DisplayColour = displayColour;
        }

        public void ParseFromString(string raw) {
            int idEnd = raw.IndexOf(' ');
            int pid;
            if (idEnd < 0) {
                throw new ArgumentException(string.Format("Malformed string: {0}", raw));
            }
            if (!int.TryParse(raw.Substring(0, idEnd), out pid)) {
                throw new ArgumentException(string.Format("Unable to parse product ID: {0}", raw.Substring(0,idEnd)));
            }
            ID = pid;
            Name = raw.Substring(idEnd+1);
        }

        public void SetColour(Color col)
        {
            DisplayColour = col;
        }

        public void AddProduct(int id)
        {
            _products.Add(id);
        }

        public void RemoveProduct(int id)
        {
            _products.Remove(id);
        }

        public List<int> GetProductIDs()
        {
            return new List<int>(_products);
        }

        public object Clone()
        {
            return new Producer(ID, Name, DisplayColour);
        }
    }
}
