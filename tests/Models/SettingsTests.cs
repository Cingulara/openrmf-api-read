using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class SettingsTests
    {
        [Fact]
        public void Test_NewSettingsIsValid()
        {
            Settings art = new Settings();
            Assert.True(art != null);
        }
    
        [Fact]
        public void Test_SettingsWithDataIsValid()
        {
            Settings set = new Settings();
            set.ConnectionString = "myConnection";
            set.Database = "user=x; database=x; password=x;";

            // test things out
            Assert.True(set != null);
            Assert.True (!string.IsNullOrEmpty(set.ConnectionString));
            Assert.True (!string.IsNullOrEmpty(set.Database));
        }
    }
}
