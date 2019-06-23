using Xunit;
using openrmf_read_api.Models;
using System;

namespace tests.Models
{
    public class ScoreTests
    {
        [Fact]
        public void Test_NewScoreIsValid()
        {
            Score score = new Score();
            Assert.True(score != null);
        }
    
        [Fact]
        public void Test_ScoreWithDataIsValid()
        {
            Score score = new Score();
            score.system = "my system";
            score.hostName = "my host name";
            score.stigRelease = "V1";
            score.stigType = "Google Chrome";
            score.created = DateTime.Now;
            score.updatedOn = DateTime.Now;


            // test things out
            Assert.True(score != null);
            Assert.True (!string.IsNullOrEmpty(score.created.ToShortDateString()));
            Assert.True (!string.IsNullOrEmpty(score.system));
            Assert.True (!string.IsNullOrEmpty(score.hostName));
            Assert.True (!string.IsNullOrEmpty(score.stigType));
            Assert.True (!string.IsNullOrEmpty(score.stigRelease));
            Assert.True (!string.IsNullOrEmpty(score.title));  // readonly from other fields
            Assert.True (score.updatedOn.HasValue);
            Assert.True (!string.IsNullOrEmpty(score.updatedOn.Value.ToShortDateString()));
            Assert.True (score.totalCat1Open == 0);
            Assert.True (score.totalCat1NotApplicable == 0);
            Assert.True (score.totalCat1NotAFinding == 0);
            Assert.True (score.totalCat1NotReviewed == 0);
            Assert.True (score.totalCat2Open == 0);
            Assert.True (score.totalCat2NotApplicable == 0);
            Assert.True (score.totalCat2NotAFinding == 0);
            Assert.True (score.totalCat2NotReviewed == 0);
            Assert.True (score.totalCat3Open == 0);
            Assert.True (score.totalCat3NotApplicable == 0);
            Assert.True (score.totalCat3NotAFinding == 0);
            Assert.True (score.totalCat3NotReviewed == 0);
            Assert.True (score.totalOpen == 0);
            Assert.True (score.totalNotApplicable == 0);
            Assert.True (score.totalNotAFinding == 0);
            Assert.True (score.totalNotReviewed == 0);
            Assert.True (score.totalCat1 == 0);
            Assert.True (score.totalCat2 == 0);
            Assert.True (score.totalCat3 == 0);
        }

        public void Test_ScoreWithCalculatedTotalsIsValid()
        {
            Score score = new Score();
            score.system = "my system";
            score.hostName = "my host name";
            score.stigRelease = "V1";
            score.stigType = "Google Chrome";
            score.created = DateTime.Now;
            score.updatedOn = DateTime.Now;


            // test things out
            Assert.True(score != null);
            Assert.True (score.totalCat1Open == 1);
            Assert.True (score.totalCat1NotApplicable == 1);
            Assert.True (score.totalCat1NotAFinding == 1);
            Assert.True (score.totalCat1NotReviewed == 1);
            Assert.True (score.totalCat2Open == 3);
            Assert.True (score.totalCat2NotApplicable == 5);
            Assert.True (score.totalCat2NotAFinding == 10);
            Assert.True (score.totalCat2NotReviewed == 20);
            Assert.True (score.totalCat3Open == 8);
            Assert.True (score.totalCat3NotApplicable == 7);
            Assert.True (score.totalCat3NotAFinding == 10);
            Assert.True (score.totalCat3NotReviewed == 10);
            Assert.True (score.totalOpen == 12);
            Assert.True (score.totalNotApplicable == 13);
            Assert.True (score.totalNotAFinding == 21);
            Assert.True (score.totalNotReviewed == 31);
            Assert.True (score.totalCat1 == 4);
            Assert.True (score.totalCat2 == 38);
            Assert.True (score.totalCat3 == 35);
        }
    }
}
