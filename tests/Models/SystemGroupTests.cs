using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class SystemGroupTests
    {
        [Fact]
        public void Test_NewSystemGroupIsValid()
        {
            SystemGroup sys = new SystemGroup();
            Assert.True(sys != null);
            Assert.True (sys.numberOfChecklists == 0);
            Assert.False (sys.updatedOn.HasValue);
            Assert.False (sys.lastComplianceCheck.HasValue);
        }
    
        [Fact]
        public void Test_SystemGroupWithDataIsValid()
        {
            SystemGroup sys = new SystemGroup();
            sys.created = DateTime.Now;
            sys.title = "My System Title";
            sys.description = "This is my System description for all items.";
            sys.numberOfChecklists = 3;
            sys.nessusFilename = "myfileservers.nessus";
            sys.rawNessusFile = "<xml></xml>";
            sys.updatedOn = DateTime.Now;
            sys.lastComplianceCheck = DateTime.Now;
            // test things out
            Assert.True(sys != null);
            Assert.True (!string.IsNullOrEmpty(sys.created.ToShortDateString()));
            Assert.True (!string.IsNullOrEmpty(sys.title));
            Assert.True (!string.IsNullOrEmpty(sys.description));
            Assert.True (!string.IsNullOrEmpty(sys.nessusFilename));
            Assert.True (!string.IsNullOrEmpty(sys.rawNessusFile));
            Assert.True (sys.numberOfChecklists == 3);
            Assert.True (sys.updatedOn.HasValue);
            Assert.True (!string.IsNullOrEmpty(sys.updatedOn.Value.ToShortDateString()));
            Assert.True (sys.lastComplianceCheck.HasValue);
            Assert.True (!string.IsNullOrEmpty(sys.lastComplianceCheck.Value.ToShortDateString()));
        }
    }
}
