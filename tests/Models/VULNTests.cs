using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class VULNTests
    {
        [Fact]
        public void Test_NewVULNIsValid()
        {
            VULN v = new VULN();
            Assert.True(v != null);
        }
    
        [Fact]
        public void Test_VULNWithDataIsValid()
        {
            VULN v = new VULN();
            v.STATUS = "my status";
            v.FINDING_DETAILS = "my status";
            v.COMMENTS = "my status";
            v.SEVERITY_OVERRIDE = "my status";
            v.SEVERITY_JUSTIFICATION = "my status";

            // test things out
            Assert.True(v != null);
            Assert.True(v.STIG_DATA != null);
            Assert.True(v.STIG_DATA.Count == 0);
            Assert.True(!string.IsNullOrEmpty(v.STATUS));
            Assert.True(!string.IsNullOrEmpty(v.FINDING_DETAILS));
            Assert.True(!string.IsNullOrEmpty(v.COMMENTS));
            Assert.True(!string.IsNullOrEmpty(v.SEVERITY_OVERRIDE));
            Assert.True(!string.IsNullOrEmpty(v.SEVERITY_JUSTIFICATION));
        }
    }
}
