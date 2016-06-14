using EmployeeInfoGrabber;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTests.PrototypeCodeSnippets
{
    [TestClass]
    public class DataSetTest
    {
        private readonly ExcelHandler excelHndlr = new ExcelHandler();
        [TestMethod]
        public void DataSetToXls()
        {
            // Create two DataTable instances.
            DataTable table1 = new DataTable("patients");
            table1.Columns.Add("Прізвище, ім'я, по батькові фізичної особи");
            table1.Columns.Add("Місце проживання");
            table1.Columns.Add("Види діяльності");
            table1.Columns.Add("Дата державної реєстрації, дата та номер запису в Єдиному державному реєстрі про включення до Єдиного державного реєстру відомостей про фізичну особу-підприємця – у разі, коли державна реєстрація фізичної особи-підприємця була проведена до набрання чинності Законом України “Про державну реєстрацію юридичних осіб та фізичних осіб-підприємців”");
            table1.Columns.Add("Дата та номер запису про проведення державної реєстрації фізичної особи-підприємця");
            table1.Columns.Add("Місцезнаходження реєстраційної справи");
            table1.Columns.Add("Дата та номер запису про взяття та зняття з обліку, назва та ідентифікаційні коди органів статистики, Міндоходів, Пенсійного фонду України, в яких фізична особа-підприємець перебуває на обліку:");
            table1.Columns.Add("Дані органів державної статистики про основний вид економічної діяльності фізичної особи-підприємця, визначений на підставі даних державних статистичних спостережень відповідно до статистичної методології за підсумками діяльності за рік");
            table1.Columns.Add("Дані про реєстраційний номер платника єдиного внеску, клас професійного ризику виробництва платника єдиного внеску за основним видом його економічної діяльності");
            table1.Columns.Add("Термін, до якого фізична особа-підприємець перебуває на обліку в органі Міндоходів за місцем попередньої реєстрації, у разі зміни місця проживання фізичної особи-підприємця");
            table1.Columns.Add("Дані про перебування фізичної особи-підприємця в процесі припинення підприємницької діяльності, банкрутства");
            table1.Columns.Add("Прізвище, ім'я, по батькові особи, яка призначена управителем майна фізичної особи-підприємця");
            table1.Columns.Add("Дата та номер запису про державну реєстрацію припинення підприємницької діяльності фізичною особою-підприємцем, підстава для його внесення");
            table1.Columns.Add("Дата відміни державної реєстрації припинення підприємницької діяльності фізичною особою-підприємцем, підстава її внесення");
            table1.Columns.Add("Дата відкриття виконавчого провадження щодо фізичної особи - підприємця (для незавершених виконавчих проваджень)");
            table1.Columns.Add("Інформація про здійснення зв'язку з фізичною особою-підприємцем");

            table1.Rows.Add("sam", 1);
            table1.Rows.Add("mark", 2);

            DataTable table2 = new DataTable("medications");
            table2.Columns.Add("");
            table2.Columns.Add("medication");
            table2.Rows.Add(1, "atenolol");
            table2.Rows.Add(2, "amoxicillin");

            // Create a DataSet and put both tables in it.
            DataSet set = new DataSet("office");
            //set.Tables.Add(table1);
            set.Tables.Add(table2);

            // Write Data Set To Xls
            
            excelHndlr.WriteExcelFile("PrototypeDat.xlsx", set);

            Assert.Fail();
        }

        [TestMethod]
        public void ExportDataSet()
        {
            //Create an Emplyee DataTable
            DataTable employeeTable = new DataTable("Employee");
            employeeTable.Columns.Add("Employee ID");
            employeeTable.Columns.Add("Employee Name");
            employeeTable.Rows.Add("1", "ABC");
            employeeTable.Rows.Add("2", "DEF");
            employeeTable.Rows.Add("3", "PQR");
            employeeTable.Rows.Add("4", "XYZ");

            //Create a Department Table
            DataTable departmentTable = new DataTable("Department");
            departmentTable.Columns.Add("Department ID");
            departmentTable.Columns.Add("Department Name");
            departmentTable.Rows.Add("1", "IT");
            departmentTable.Rows.Add("2", "HR");
            departmentTable.Rows.Add("3", "Finance");

            //Create a DataSet with the existing DataTables
            DataSet ds = new DataSet("Organization");
            
            //ds.Tables.Add(employeeTable);
            ds.Tables.Add(departmentTable);
            excelHndlr.WriteExcelFile("testOut.xls", ds);
        }
    }
}
