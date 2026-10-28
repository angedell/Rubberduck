using System.Text;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ApiBodyEncodingTests
    {
        [TestCase(null)]
        [TestCase("")]
        [TestCase("application/json")]
        [TestCase("application/json; charset=utf-8")]
        [TestCase("application/json; charset=\"UTF-8\"")]
        [TestCase("application/json; charset=nonsense")]
        public void Json_IsUtf8_UnlessAnotherCharsetIsStated(string contentType)
        {
            Assert.AreEqual(Encoding.UTF8.CodePage, ApiServer.BodyEncoding(contentType).CodePage);
        }

        [Test]
        public void AStatedCharset_IsHonoured()
        {
            Assert.AreEqual(Encoding.Unicode.CodePage, ApiServer.BodyEncoding("application/json; charset=utf-16").CodePage);
        }
    }
}
