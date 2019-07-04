using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class ASSETTests
    {
        [Fact]
        public void Test_NewASSETIsValid()
        {
            ASSET asset = new ASSET();
            Assert.True(asset != null);
        }
    
        [Fact]
        public void Test_ASSETWithDataIsValid()
        {
            ASSET asset = new ASSET();
    		asset.ROLE  = "myRole";
    		asset.ASSET_TYPE  = "myRole";
    		asset.HOST_NAME  = "myRole";
    		asset.HOST_IP  = "myRole";
    		asset.HOST_MAC  = "myRole";
    		asset.HOST_FQDN  = "myRole";
    		asset.TECH_AREA  = "myRole";
    		asset.TARGET_KEY  = "myRole";
    		asset.WEB_OR_DATABASE  = "myRole";
    		asset.WEB_DB_SITE  = "myRole";
    		asset.WEB_DB_INSTANCE  = "myRole";
            
            // test things out
            Assert.True(asset != null);
            Assert.True(!string.IsNullOrEmpty(asset.ROLE));
            Assert.True(!string.IsNullOrEmpty(asset.ASSET_TYPE));
            Assert.True(!string.IsNullOrEmpty(asset.HOST_NAME));
            Assert.True(!string.IsNullOrEmpty(asset.HOST_IP));
            Assert.True(!string.IsNullOrEmpty(asset.HOST_MAC));
            Assert.True(!string.IsNullOrEmpty(asset.HOST_FQDN));
            Assert.True(!string.IsNullOrEmpty(asset.TECH_AREA));
            Assert.True(!string.IsNullOrEmpty(asset.TARGET_KEY));
            Assert.True(!string.IsNullOrEmpty(asset.WEB_OR_DATABASE));
            Assert.True(!string.IsNullOrEmpty(asset.WEB_DB_SITE));
            Assert.True(!string.IsNullOrEmpty(asset.WEB_DB_INSTANCE));
        }
    }
}
