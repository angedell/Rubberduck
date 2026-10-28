using System;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.PackageResources;
using RubberduckTests.PackageResources;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ApiResourceTests
    {
        private class Files : IProjectFiles
        {
            public string Document;

            public string DocumentOf(string project)
                => project == "VBAProject" || project == "p1" ? Document : throw ApiException.NotFound($"No open project '{project}'.");
        }

        private string _folder;
        private string _document;
        private int _asked;
        private bool _allow;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _folder = Path.Combine(Path.GetTempPath(), "rd-api-resources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
            _document = Path.Combine(_folder, "Book.xlsm");
            PackageResourceServiceTests.WriteDocument(_document, withRibbon: true);
            _asked = 0;
            _allow = true;

            var approval = new ApiWriteApproval(_ => { _asked++; return _allow; });
            var service = new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"));
            var resources = new ResourceApi(new Files { Document = _document }, service, approval);
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, resources);
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                Directory.Delete(_folder, true);
            }
            catch (IOException)
            {
            }
        }

        private ApiResponse Send(string method, string path, object body = null, string ifMatch = null)
            => _router.Handle(new ApiRequest { Method = method, Path = path, Body = body == null ? null : JObject.FromObject(body).ToString(), IfMatch = ifMatch, Client = "test" });

        private static string ErrorCode(ApiResponse response) => (string)JObject.FromObject(response.Body)["error"]["code"];
        private static JObject Json(ApiResponse response) => JObject.FromObject(response.Body);

        private static string PngBase64 => Convert.ToBase64String(PackageResourceServiceTests.Png);

        [Test]
        public void List_ShowsTheRibbonAndThePictures_WithoutAskingForApproval()
        {
            var response = Send("GET", "/api/v1/projects/p1/resources");

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.IsTrue((bool)json["Supported"]);
            CollectionAssert.AreEqual(new[] { "customUI14.xml", "logo.png" }, json["Resources"].Select(r => (string)r["Name"]).ToArray());
            Assert.AreEqual("ribbon", (string)json["Resources"][0]["Kind"]);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void List_ForAnUnsavedProject_SaysWhyThereIsNothing()
        {
            File.Delete(_document);

            var json = Json(Send("GET", "/api/v1/projects/p1/resources"));

            Assert.IsFalse((bool)json["Supported"]);
            StringAssert.Contains("Open XML", (string)json["Note"]);
        }

        [Test]
        public void Read_GivesTheXmlAndItsHash_AndAPictureAsBase64()
        {
            var xml = Send("GET", "/api/v1/projects/p1/resources/customUI14.xml");
            var picture = Send("GET", "/api/v1/projects/p1/resources/logo.png");

            StringAssert.Contains("<ribbon", (string)Json(xml)["Xml"]);
            Assert.AreEqual(xml.ETag, (string)Json(xml)["Hash"]);
            Assert.AreEqual(PngBase64, (string)Json(picture)["ContentBase64"]);
        }

        [Test]
        public void Read_OfAnUnknownResource_Is404()
        {
            Assert.AreEqual(404, Send("GET", "/api/v1/projects/p1/resources/nope.png").Status);
            Assert.AreEqual(404, Send("GET", "/api/v1/projects/nope/resources").Status);
        }

        [Test]
        public void WriteXml_IsStaged_ApprovedOnce_AndShowsInTheNextRead()
        {
            var hash = Send("GET", "/api/v1/projects/p1/resources/customUI14.xml").ETag;
            var edited = PackageResourceServiceTests.Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>");

            var response = Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { xml = edited, expectedHash = hash });
            var again = Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { xml = edited, expectedHash = Json(Send("GET", "/api/v1/projects/p1/resources/customUI14.xml"))["Hash"] });

            Assert.AreEqual(200, response.Status);
            StringAssert.Contains("closed", (string)Json(response)["Note"]);
            Assert.AreEqual("modified", (string)Json(response)["Resource"]["Pending"]);
            Assert.AreEqual(200, again.Status);
            Assert.AreEqual(1, _asked);
            Assert.AreEqual(edited, (string)Json(Send("GET", "/api/v1/projects/p1/resources/customUI14.xml"))["Xml"]);
            Assert.IsTrue((bool)Json(Send("GET", "/api/v1/projects/p1/resources"))["PendingChanges"]);
        }

        [Test]
        public void WriteXml_WithoutOrWithAStaleHash_IsRefused()
        {
            var edited = PackageResourceServiceTests.Ribbon;

            Assert.AreEqual("hash_required", ErrorCode(Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { xml = edited })));
            Assert.AreEqual("stale_hash", ErrorCode(Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { xml = edited, expectedHash = "0000" })));
        }

        [Test]
        public void WriteXml_ThatOfficeWouldIgnore_Is400_WithTheLine()
        {
            var hash = Send("GET", "/api/v1/projects/p1/resources/customUI14.xml").ETag;

            var response = Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { xml = "<customUI\r\n<", expectedHash = hash });

            Assert.AreEqual(400, response.Status);
            StringAssert.Contains("Line", (string)Json(response)["error"]["message"]);
        }

        [Test]
        public void WriteAPicture_NeedsNoHashWhenNew()
        {
            var response = Send("PUT", "/api/v1/projects/p1/resources/new.png", new { contentBase64 = PngBase64 });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("added", (string)Json(response)["Resource"]["Pending"]);
        }

        [Test]
        public void Write_MixingUpXmlAndPictures_Is400_BeforeAnyApprovalIsAsked()
        {
            Assert.AreEqual(400, Send("PUT", "/api/v1/projects/p1/resources/new.png", new { xml = "<customUI/>" }).Status);
            Assert.AreEqual(400, Send("PUT", "/api/v1/projects/p1/resources/customUI14.xml", new { contentBase64 = PngBase64 }).Status);
            Assert.AreEqual(400, Send("PUT", "/api/v1/projects/p1/resources/new.png", new { contentBase64 = "!!!" }).Status);
            Assert.AreEqual(400, Send("PUT", "/api/v1/projects/p1/resources/new.png", new { }).Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Write_WhenTheUserSaysNo_Is403()
        {
            _allow = false;

            var response = Send("PUT", "/api/v1/projects/p1/resources/new.png", new { contentBase64 = PngBase64 });

            Assert.AreEqual(403, response.Status);
            Assert.AreEqual("write_denied", ErrorCode(response));
        }

        [Test]
        public void Delete_StagesTheRemoval_AndNeedsTheHash()
        {
            var hash = Send("GET", "/api/v1/projects/p1/resources/logo.png").ETag;

            Assert.AreEqual(428, Send("DELETE", "/api/v1/projects/p1/resources/logo.png").Status);
            Assert.AreEqual(412, Send("DELETE", "/api/v1/projects/p1/resources/logo.png", null, "\"0000\"").Status);
            var response = Send("DELETE", "/api/v1/projects/p1/resources/logo.png", null, "\"" + hash + "\"");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("removed", (string)Json(response)["Resource"]["Pending"]);
            Assert.AreEqual("removed", (string)Json(Send("GET", "/api/v1/projects/p1/resources"))["Resources"].Single(r => (string)r["Name"] == "logo.png")["Pending"]);
        }

        [Test]
        public void DeletingTheCollection_DiscardsWhatIsStaged_AndLeavesTheFileAlone()
        {
            var hash = Send("GET", "/api/v1/projects/p1/resources/logo.png").ETag;
            Send("DELETE", "/api/v1/projects/p1/resources/logo.png", null, "\"" + hash + "\"");
            Assert.IsTrue((bool)Json(Send("GET", "/api/v1/projects/p1/resources"))["PendingChanges"]);

            var response = Send("DELETE", "/api/v1/projects/p1/resources");

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.IsFalse((bool)json["PendingChanges"]);
            Assert.AreEqual(JTokenType.Null, json["Resources"].Single(r => (string)r["Name"] == "logo.png")["Pending"].Type);
        }

        [Test]
        public void DeletingTheCollection_WithNothingStaged_AsksNobody()
        {
            var before = _asked;

            Assert.AreEqual(200, Send("DELETE", "/api/v1/projects/p1/resources").Status);
            Assert.AreEqual(before, _asked);
        }

        [Test]
        public void Delete_OfRibbonXml_Is400()
        {
            var hash = Send("GET", "/api/v1/projects/p1/resources/customUI14.xml").ETag;

            Assert.AreEqual(400, Send("DELETE", "/api/v1/projects/p1/resources/customUI14.xml", null, "\"" + hash + "\"").Status);
        }

        [Test]
        public void WithoutResourceSupport_TheRouteDoesNotExist()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            var response = router.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/resources", Client = "test" });

            Assert.AreEqual(404, response.Status);
        }

        [Test]
        public void Mcp_OffersTheResourceTools_AndRunsThem()
        {
            JObject Call(object message)
                => Json(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = JObject.FromObject(message).ToString(), Client = "test" }));

            var tools = Call(new { jsonrpc = "2.0", id = 1, method = "tools/list" })["result"]["tools"].Select(t => (string)t["name"]).ToList();
            var listed = Call(new { jsonrpc = "2.0", id = 2, method = "tools/call", @params = new { name = "list_resources", arguments = new { project = "p1" } } });

            CollectionAssert.IsSubsetOf(new[] { "list_resources", "read_resource", "write_resource", "delete_resource" }, tools);
            Assert.IsFalse((bool)listed["result"]["isError"]);
            StringAssert.Contains("customUI14.xml", (string)listed["result"]["content"][0]["text"]);
        }
    }
}
