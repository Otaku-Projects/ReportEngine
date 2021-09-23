using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JasperReport.ReportDataModel
{
    public class HitRateDataModel
    {
        private string _id;
        private string _office;
        private string _product;
        private string _city;
        private int _numOfDesign;
        private int _numOfContracted;
        private decimal _designHitRate;
        private int _numOfColorWays;
        private int _numOfItems;
        private decimal _colorwayHitRate;

        public string Id { get => _id; set => _id = value; }
        public string Office { get => _office; set => _office = value; }
        public string Product { get => _product; set => _product = value; }
        public string City { get => _city; set => _city = value; }
        public int NumOfDesign { get => _numOfDesign; set => _numOfDesign = value; }
        public int NumOfContracted { get => _numOfContracted; set => _numOfContracted = value; }
        public decimal DesignHitRate { get => _designHitRate; set => _designHitRate = value; }
        public int NumOfColorWays { get => _numOfColorWays; set => _numOfColorWays = value; }
        public int NumOfItems { get => _numOfItems; set => _numOfItems = value; }
        public decimal ColorwayHitRate { get => _colorwayHitRate; set => _colorwayHitRate = value; }

        public HitRateDataModel() { }

        public HitRateDataModel(
            string id
            , string office
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
            this._product  = product;
            this._city = city;
            this._numOfDesign = numOfDesign;
            this._numOfContracted = numOfContracted;
            this._designHitRate = designHitRate;
            this._numOfColorWays = numOfColorWays;
            this._numOfItems = numOfItems;
            this._colorwayHitRate = colorwayHitRate;
        }

        //public HitRateDataModel(DataSet _dataSet)
        //{
        //    if (_dataSet == null) throw new NoNullAllowedException();
        //    if (_dataSet.Tables.Count == 0) throw new NoNullAllowedException();

        //    this.dataSet = _dataSet;
        //    this.CreateDummyData1("GeneralView");
        //}

        //public DataTable AddRowToHitRateDataTable(DataTable _dataTable, int _count)
        //{
        //    DataRow _dRow = _dataTable.NewRow();

        //    for (int i = 0; i < _count; i++)
        //    {
        //        _dRow = _dataTable.NewRow();
        //        _dRow[0] = i;
        //        _dRow[1] = Faker.Company.Name();
        //        _dRow[2] = Faker.Company.Name();
        //        _dRow[3] = Faker.Address.City();
        //        _dRow[4] = Faker.RandomNumber.Next(1, 5);
        //        _dRow[5] = Faker.RandomNumber.Next(1, 30);
        //        _dRow[6] = Faker.RandomNumber.Next((long)0, (long)1);
        //        _dRow[7] = Faker.RandomNumber.Next(1, 5);
        //        _dRow[8] = Faker.RandomNumber.Next(50, 10000);
        //        _dRow[9] = Faker.RandomNumber.Next((long)0, (long)1);

        //        _dataTable.Rows.Add(_dRow);
        //    }


        //    return _dataTable;
        //}

        //public DataTable AddDataColumnToDataTable(DataTable _dataTable)
        //{
        //    List<DataColumn> _dataColumnList = new List<DataColumn>();
        //    DataColumn _dataColumn1;
        //    _dataColumn1 = this.CreateDataColumn("id", typeof(int));
        //    _dataColumnList.Add(_dataColumn1);

        //    _dataColumn1 = this.CreateDataColumn("office", typeof(string));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

        //    _dataColumn1 = this.CreateDataColumn("product", typeof(string));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

        //    _dataColumn1 = this.CreateDataColumn("city", typeof(string));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

        //    _dataColumn1 = this.CreateDataColumn("numOfDesign", typeof(int));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
        //    _dataColumn1 = this.CreateDataColumn("numOfContracted", typeof(int));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
        //    _dataColumn1 = this.CreateDataColumn("designHitRate", typeof(Decimal));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
        //    _dataColumn1 = this.CreateDataColumn("numOfColorWays", typeof(int));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
        //    _dataColumn1 = this.CreateDataColumn("numOfItems", typeof(int));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
        //    _dataColumn1 = this.CreateDataColumn("colorwayHitRate", typeof(Decimal));
        //    _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

        //    //DataColumn a = new DataColumn()

        //    foreach (var _col in _dataColumnList)
        //    {
        //        _dataTable.Columns.Add(_col);
        //    }
        //    return _dataTable;
        //}

        //public DataColumn CreateDataColumn(string _columnName, Type _dataType)
        //{
        //    DataColumn _dataCol = new DataColumn(_columnName, _dataType);
        //    return _dataCol;
        //}

        //public void CreateDummyData1(string _tableName)
        //{
        //    if (string.IsNullOrEmpty(_tableName))
        //    {
        //        Guid obj = Guid.NewGuid();
        //        _tableName = obj.ToString();
        //    }
        //    DataTable _table = new DataTable(_tableName);
        //    //_table = this.AddRowToHitRateDataTable(_table.Copy(), 100);
        //    this.AddDataColumnToDataTable(_table);
        //    this.AddRowToHitRateDataTable(_table, 100);

        //    this.dataSet.Tables.Add(_table);

        //    //return _table;
        //}

        //public DataSet GetDataSet()
        //{
        //    return this.dataSet;
        //}

        //public void SetDataSet(DataSet _dataSet)
        //{
        //    this.dataSet = _dataSet;
        //}
    }
}
