﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace CoreSystemConsole.ReportDataModel
{
    class HitRateDataView
    {
        private DataSet dataSet;
        private IDictionary<string, object> dataSetObj;
        public HitRateDataView()
        {
            DataSet _dataSet = new DataSet();
            this.dataSet = _dataSet;

            this.dataSetObj = new ExpandoObject() as IDictionary<string, object>;
        }
        public HitRateDataView(DataSet _dataSet)
        {
            if (_dataSet == null) throw new NoNullAllowedException();
            if (_dataSet.Tables.Count == 0) throw new NoNullAllowedException();

            this.dataSet = _dataSet;

            this.dataSetObj = new ExpandoObject() as IDictionary<string, object>;
        }

        public DataTable AddRowToHitRateDataTable(DataTable _dataTable, int _maxLimit)
        {
            DataRow _dRow = _dataTable.NewRow();

            List<string> randomStrings = Enumerable.Range(1, Convert.ToInt32(Math.Round(_maxLimit / 2.0m)))
                       .Select(_ => Faker.Company.Name())
                       .ToList();
            for (int i = 0; i < _maxLimit; i++)
            {
                _dRow = _dataTable.NewRow();
                _dRow[0] = i;
                _dRow[1] = randomStrings[Convert.ToInt32(i % Math.Round(_maxLimit / 2.0m))];
                _dRow[2] = Faker.Company.Name();
                _dRow[3] = Faker.Internet.DomainWord();
                _dRow[4] = Faker.Address.City();
                _dRow[5] = Faker.RandomNumber.Next(1, 100);
                _dRow[6] = Faker.RandomNumber.Next(0, 100);
                _dRow[7] = Faker.RandomNumber.Next((long)0, (long)1);
                _dRow[8] = Faker.RandomNumber.Next(1, 5);
                _dRow[9] = Faker.RandomNumber.Next(50, 1000);
                _dRow[10] = Faker.RandomNumber.Next((long)0, (long)1);

                _dataTable.Rows.Add(_dRow);
            }

            return _dataTable;
        }

        public DataTable AddDataColumnToDataTable(DataTable _dataTable)
        {
            List<DataColumn> _dataColumnList = new List<DataColumn>();
            DataColumn _dataColumn1;

            var dataModelInstance = new HitRateDataModel2();
            foreach (PropertyInfo propertyInfo in dataModelInstance.GetType().GetProperties())
            {
                _dataColumn1 = this.CreateDataColumn(propertyInfo.Name, propertyInfo.PropertyType);
                _dataColumnList.Add(_dataColumn1);
            }

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

        public void CreateDummyData1()
        {
            this.dataSetObj = new ExpandoObject() as IDictionary<string, object>;

            this.dataSetObj.Add("number", "0123456789~!@#$%^&*()_+");
            this.dataSetObj.Add("staffname", "Peter Pan (sys0999)");

            this.CreateDummyDataGeneralView1();
            this.CreateDummyDataSeller();
            this.CreateDummyDataBuyer();
        }

        public void CreateDummyData2()
        {
            this.dataSetObj = new ExpandoObject() as IDictionary<string, object>;

            this.dataSetObj.Add("number", "0123456789~!@#$%^&*()_+");
            this.dataSetObj.Add("staffname", "Peter Pan (sys0999)");

            this.CreateDummyDataGeneralView2();
            this.CreateDummyDataSeller();
            this.CreateDummyDataBuyer();
        }
        public void CreateDummyData3()
        {
            this.dataSetObj = new ExpandoObject() as IDictionary<string, object>;

            this.dataSetObj.Add("number", "0123456789~!@#$%^&*()_+");
            this.dataSetObj.Add("staffname", "Peter Pan (sys0999)");

            this.CreateDummy3_HitRateData();
            this.CreateDummy3_StaffProfile();
            this.CreateDummyDataSeller();
            this.CreateDummyDataBuyer();
        }

        public void CreateDummyDataGeneralView1(string _tableName = "GeneralView")
        {
            dynamic _obj = new ExpandoObject();
            _obj = new[]
            {
                new { name = "NextCore System", price = 100000 },
                new { name = "Implementation (3 man-days)", price = 20000 },
                new { name = "Annual Support (24 man-hours)", price = 30000 },
                new { name = "Volume License", price = 4000 }
            };

            this.dataSetObj.Add(_tableName, _obj);
        }

        public void CreateDummyDataGeneralView2(string _tableName = "GeneralView")
        {
            List<dynamic> _obj = new List<dynamic>();
            for (int i = 0; i < 100; i++)
            {
                _obj.Add(new
                {
                    name = Faker.Company.Name(),
                    price = Faker.RandomNumber.Next(100, 1000000),
                });
            }
            this.dataSetObj.Add(_tableName, _obj);
        }

        public void CreateDummy3_HitRateData()
        {
            List<dynamic> _obj = new List<dynamic>();
            string _tableName = "GeneralView";
            string officeName = string.Empty;
            officeName = Faker.Address.Country();

            // generate records count
            int maxLimit = 10;

            // create datatable
            DataTable _table = new DataTable(_tableName);
            this.AddDataColumnToDataTable(_table);
            this.AddRowToHitRateDataTable(_table, maxLimit);
            this.dataSet.Tables.Add(_table);

            // create datasetObj
            List<string> randomStrings = Enumerable.Range(1, maxLimit)
                       .Select(_ => Faker.Company.Name())
                       .ToList();
            for (int i = 0; i < maxLimit; i++)
            {
                //_obj.Add(new
                //{
                //    OfficeName = officeName,
                //    Department = Faker.Company.Name(),
                //    ProductTeam = Faker.Internet.DomainWord(),
                //    City = Faker.Address.City(),
                //    NumOfDesign = Faker.RandomNumber.Next(1, 100),
                //    NumOfDesignContracted = Faker.RandomNumber.Next(0, 100),
                //    DesignHitRate = Faker.RandomNumber.Next(0, 100),
                //    NumOfColorways = Faker.RandomNumber.Next(1, 100),
                //    NumOfItem = Faker.RandomNumber.Next(0, 100),
                //    ColorwayHitRate = Faker.RandomNumber.Next(0, 100),
                //});

                _obj.Add(new
                {
                    OfficeName = randomStrings[Convert.ToInt32(i % Math.Round(maxLimit / 2.0m))],
                    Department = Faker.Company.Name(),
                    ProductTeam = Faker.Internet.DomainWord(),
                    City = Faker.Address.City(),
                    NumOfDesign = Faker.RandomNumber.Next(1, 100),
                    NumOfDesignContracted = Faker.RandomNumber.Next(0, 100),
                    DesignHitRate = Faker.RandomNumber.Next(0, 100),
                    NumOfColorways = Faker.RandomNumber.Next(1, 100),
                    NumOfItems = Faker.RandomNumber.Next(0, 100),
                    ColorwayHitRate = Faker.RandomNumber.Next(0, 100),
                });
            }
            this.dataSetObj.Add(_tableName, _obj);
        }

        public void CreateDummy3_StaffProfile()
        {
            List<dynamic> _obj = new List<dynamic>();
            string _tableName = "StaffView";
            for (int i = 0; i < 100; i++)
            {
                Random gen = new Random();
                DateTime dateOfBirth = new DateTime(1960, 1, 1);
                int lifeDay = (DateTime.Today - dateOfBirth).Days;
                dateOfBirth = dateOfBirth.AddDays(gen.Next(lifeDay));

                DateTime employmentDate = new DateTime(1990, 1, 1);
                int hiringDays = (DateTime.Today - employmentDate).Days;
                employmentDate = employmentDate.AddDays(gen.Next(hiringDays));

                Random random = new Random();
                var genderNumber = random.Next(0, 2);

                _obj.Add(new
                {
                    UUIID = Faker.RandomNumber.Next(10001, 99999),
                    StaffID = (i+1),
                    Address = Faker.Address.Country()+" "+ Faker.Address.City()+ Faker.Address.StreetName(),
                    FirstName = Faker.Name.First(),
                    LastName = Faker.Name.Last(),
                    DateOfBirth = dateOfBirth,
                    Gender = (genderNumber==0) ? "M" : "F",
                    PassportID = Faker.Lorem.Words(7),
                    EmploymentDate = employmentDate,
                });
            }
            this.dataSetObj.Add(_tableName, _obj);
        }

        public void CreateDummyDataSeller(string _tableName = "seller")
        {
            dynamic _obj = new ExpandoObject();
            _obj = new
            {
                name = "Next Step Webs, Inc.",
                road = "12345 Sunny Road",
                country = "Sunnyville, TX 12345"
            };

            this.dataSetObj.Add(_tableName, _obj);
        }

        public void CreateDummyDataBuyer(string _tableName = "buyer")
        {
            dynamic _obj = new ExpandoObject();
            _obj = new
            {
                name = "Acme Corp.",
                road = "16 Johnson Road",
                country = "Paris, France 8060"
            };

            this.dataSetObj.Add(_tableName, _obj);
        }

        public DataSet GetDataSet()
        {
            return this.dataSet;
        }

        public void SetDataSet(DataSet _dataSet)
        {
            this.dataSet = _dataSet;
        }

        public IDictionary<string, object> GetDataSetObj()
        {
            return this.dataSetObj;
        }

        public void SetDataSetObj(IDictionary<string, object> _dataSetObj)
        {
            this.dataSetObj = _dataSetObj;
        }
    }
}
