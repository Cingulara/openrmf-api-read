using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class iSTIGTests
    {
        [Fact]
        public void Test_NewiSTIGIsValid()
        {
            iSTIG iStig = new iSTIG();
            Assert.True(iStig != null);
        }
    
        [Fact]
        public void Test_iSTIGWithDataIsValid()
        {
            iSTIG iStig = new iSTIG();
            // test things out
            Assert.True(iStig != null);
            Assert.True(iStig.STIG_INFO != null);
            Assert.True(iStig.VULN != null);
            Assert.True(iStig.VULN.Count == 0);
        }
    }
}
