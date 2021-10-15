using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ReportDataModel
{
    public class HitRateDataModel2
    {
        private string _id;
        private string _office;
        private string _product;
        private string _department;
        private string _city;
        private int _numOfDesign;
        private int _numOfContracted;
        private decimal _designHitRate;
        private int _numOfColorWays;
        private int _numOfItems;
        private decimal _colorwayHitRate;

        public string Id { get => _id; set => _id = value; }
        public string OfficeName { get => _office; set => _office = value; }
        public string Department { get => _department; set => _department = value; }
        public string Product { get => _product; set => _product = value; }
        public string City { get => _city; set => _city = value; }
        public int NumOfDesign { get => _numOfDesign; set => _numOfDesign = value; }
        public int NumOfDesignContracted { get => _numOfContracted; set => _numOfContracted = value; }
        public decimal DesignHitRate { get => _designHitRate; set => _designHitRate = value; }
        public int NumOfColorways { get => _numOfColorWays; set => _numOfColorWays = value; }
        public int NumOfItems { get => _numOfItems; set => _numOfItems = value; }
        public decimal ColorwayHitRate { get => _colorwayHitRate; set => _colorwayHitRate = value; }

        public HitRateDataModel2() { }

        public HitRateDataModel2(
            string id
            , string office
            , string department
            , string product
            , string city
            ,int numOfDesign
            , int numOfContracted
            , decimal designHitRate
            , int numOfColorWays
            , int numOfItems
            , decimal colorwayHitRate)
        {
            this._id = id;
            this._office = office;
            this._department = department;
            this._product = product;
            this._city = city;
            this._numOfDesign = numOfDesign;
            this._numOfContracted = numOfContracted;
            this._designHitRate = designHitRate;
            this._numOfColorWays = numOfColorWays;
            this._numOfItems = numOfItems;
            this._colorwayHitRate = colorwayHitRate;
        }
    }
}
