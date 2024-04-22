using Apchy.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MSTesting4
{
    [TestClass]
    public class DateTimeHelperTesting
    {
        [TestMethod]

        public void ShouldAlmostEqualNow()
        {
            DateTime dt = DateTime.Now;

            DateTimeHelper.RegisterServerTime(dt);

            Thread.Sleep(100);

            var serverNow = DateTimeHelper.ServerNow();
            var now = DateTime.Now;

            var ts = now - serverNow;

            var differ = ts.TotalMilliseconds;

            Assert.IsTrue(differ < 10);
        }
    }
}
