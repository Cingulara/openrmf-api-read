using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class STIG_INFOTests
    {
        [Fact]
        public void Test_NewSTIG_INFOIsValid()
        {
            STIG_INFO data = new STIG_INFO();
            Assert.True(data != null);
        }
    
        [Fact]
        public void Test_STIG_INFOWithDataIsValid()
        {
            STIG_INFO data = new STIG_INFO();

            // test things out
            Assert.True(data != null);
            Assert.True(data.SI_DATA != null);
            Assert.True(data.SI_DATA.Count == 0);
        }
    }
}
