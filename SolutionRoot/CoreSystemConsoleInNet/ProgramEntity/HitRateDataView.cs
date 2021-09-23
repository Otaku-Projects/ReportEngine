using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsoleInNet.ProgramEntity
{
    class HitRateDataView
    {
        private DataSet dataSet;
        public HitRateDataView()
        {
            DataSet _dataSet = new DataSet();
            //_dataSet.Tables.Add("GeneralView");
            this.dataSet = _dataSet;
            this.CreateDummyData1("GeneralView");
        }
        public HitRateDataView(DataSet _dataSet)
        {
            if (_dataSet == null) throw new NoNullAllowedException();
            if (_dataSet.Tables.Count == 0) throw new NoNullAllowedException();

            this.dataSet = _dataSet;
            this.CreateDummyData1("GeneralView");
        }

        public DataTable AddRowToHitRateDataTable(DataTable _dataTable, int _count)
        {
            DataRow _dRow = _dataTable.NewRow();

            for (int i = 0; i < _count; i++)
            {
                _dRow = _dataTable.NewRow();
                _dRow[0] = i;
                _dRow[1] = Faker.Company.Name();
                _dRow[2] = Faker.Company.Name();
                _dRow[3] = Faker.Address.City();
                _dRow[4] = Faker.RandomNumber.Next(1, 5);
                _dRow[5] = Faker.RandomNumber.Next(1, 30);
                _dRow[6] = Faker.RandomNumber.Next((long)0, (long)1);
                _dRow[7] = Faker.RandomNumber.Next(1, 5);
                _dRow[8] = Faker.RandomNumber.Next(50, 10000);
                _dRow[9] = Faker.RandomNumber.Next((long)0, (long)1);

                _dataTable.Rows.Add(_dRow);
            }


            return _dataTable;
        }

        public DataTable AddDataColumnToDataTable(DataTable _dataTable)
        {
            List<DataColumn> _dataColumnList = new List<DataColumn>();
            DataColumn _dataColumn1;
            _dataColumn1 = this.CreateDataColumn("id", typeof(int));
            _dataColumnList.Add(_dataColumn1);

            _dataColumn1 = this.CreateDataColumn("office", typeof(string));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

            _dataColumn1 = this.CreateDataColumn("product", typeof(string));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

            _dataColumn1 = this.CreateDataColumn("city", typeof(string));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

            _dataColumn1 = this.CreateDataColumn("numOfDesign", typeof(int));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
            _dataColumn1 = this.CreateDataColumn("numOfContracted", typeof(int));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
            _dataColumn1 = this.CreateDataColumn("designHitRate", typeof(Decimal));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
            _dataColumn1 = this.CreateDataColumn("numOfColorWays", typeof(int));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
            _dataColumn1 = this.CreateDataColumn("numOfItems", typeof(int));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));
            _dataColumn1 = this.CreateDataColumn("colorwayHitRate", typeof(Decimal));
            _dataColumnList.Add(new DataColumn(_dataColumn1.ColumnName, _dataColumn1.DataType));

            //DataColumn a = new DataColumn()

            foreach (var _col in _dataColumnList)
            {
                _dataTable.Columns.Add(_col);
            }
            return _dataTable;
        }

        public DataColumn CreateDataColumn(string _columnName, Type _dataType)
        {
            DataColumn _dataCol = new DataColumn(_columnName, _dataType);
            return _dataCol;
        }

        public void CreateDummyData1(string _tableName)
        {
            if (string.IsNullOrEmpty(_tableName))
            {
                Guid obj = Guid.NewGuid();
                _tableName = obj.ToString();
            }
            DataTable _table = new DataTable(_tableName);
            //_table = this.AddRowToHitRateDataTable(_table.Copy(), 100);
            this.AddDataColumnToDataTable(_table);
            this.AddRowToHitRateDataTable(_table, 100);

            this.dataSet.Tables.Add(_table);

            //return _table;
        }

        public DataSet GetDataSet()
        {
            return this.dataSet;
        }

        public void SetDataSet(DataSet _dataSet)
        {
            this.dataSet = _dataSet;
        }
    }
}
