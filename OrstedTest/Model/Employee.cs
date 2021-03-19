using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OrstedTest.Model
{
    public class Employee:IDisposable
    {
        public int EmployeeNo { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Status { get; set; }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
