using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.AspNetCore.Mvc;
using OrstedTest.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace OrstedTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReadExcelController : ControllerBase
    {   
        [HttpPost]
        [Route("getEmployee")]
        public ActionResult<IEnumerable<Employee>> ReadFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"../Excel.xlsx");
            Excel._Worksheet xlWorksheet =(Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Sheets["Sheet1"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<Employee> employeeList = new List<Employee>();

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                Employee emp = new Employee();

                emp.EmployeeNo = Convert.ToInt32(xlRange.Cells[i, 1].ToString());
                emp.FirstName = xlRange.Cells[i, 2].ToString();
                emp.LastName = xlRange.Cells[i, 3].ToString();
                emp.Status = xlRange.Cells[i, 4].ToString();
                employeeList.Add(emp);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return employeeList;
        }
    }
}