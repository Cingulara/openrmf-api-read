using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class SI_DATATests
    {
        [Fact]
        public void Test_NewSI_DATAIsValid()
        {
            SI_DATA data = new SI_DATA();
            Assert.True(data != null);
        }
    
        [Fact]
        public void Test_SI_DATAWithDataIsValid()
        {
            SI_DATA data = new SI_DATA();
            data.SID_DATA = "mydata";
            data.SID_NAME = "myName";

            // test things out
            Assert.True(data != null);
            Assert.True(!string.IsNullOrEmpty(data.SID_DATA));
            Assert.True(!string.IsNullOrEmpty(data.SID_NAME));
        }
    }
}
