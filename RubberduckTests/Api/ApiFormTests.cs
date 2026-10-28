using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ApiFormTests
    {
        private class FakeForms : IFormControlStore
        {
            public List<string> Calls = new List<string>();
            public string ProgId;
            public JObject Properties;

            public JObject ListControls(string project, string module)
            {
                Calls.Add("List");
                return new JObject { ["controls"] = new JArray(new JObject { ["name"] = "txt", ["type"] = "TextBox" }) };
            }

            public JObject AddControl(string project, string module, string progId, string name, JObject properties)
            {
                Calls.Add("Add");
                ProgId = progId;
                Properties = properties;
                return new JObject { ["name"] = name };
            }

            public JObject SetControl(string project, string module, string name, JObject properties)
            {
                Calls.Add("Set");
                Properties = properties;
                return new JObject { ["name"] = name };
            }

            public void DeleteControl(string project, string module, string name) => Calls.Add("Delete");
        }

        private FakeForms _store;
        private int _asked;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _store = new FakeForms();
            _asked = 0;
            var approval = new ApiWriteApproval(_ => { _asked++; return true; });
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, forms: new FormApi(_store, approval));
        }

        private ApiResponse Send(string method, string path, object body = null)
            => _router.Handle(new ApiRequest { Method = method, Path = "/api/v1" + path, Body = body == null ? null : JObject.FromObject(body).ToString(), Client = "agent" });

        [Test]
        public void List_NeedsNoApproval()
        {
            var response = Send("GET", "/projects/p/modules/frm/controls");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("txt", (string)JObject.FromObject(response.Body)["controls"][0]["name"]);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Add_ShortTypeBecomesProgId_AndFieldsBecomeProperties()
        {
            var response = Send("POST", "/projects/p/modules/frm/controls", new { type = "combobox", name = "cmbStatus", left = 12, top = 12, width = 200, height = 20, properties = new { Style = 2 } });

            Assert.AreEqual(201, response.Status);
            Assert.AreEqual("Forms.ComboBox.1", _store.ProgId);
            Assert.AreEqual(12, (int)_store.Properties["Left"]);
            Assert.AreEqual(200, (int)_store.Properties["Width"]);
            Assert.AreEqual(2, (int)_store.Properties["Style"]);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Add_ProgIdIsAccepted_ButOtherControlsAreNot()
        {
            Assert.AreEqual(201, Send("POST", "/projects/p/modules/frm/controls", new { type = "Forms.TextBox.1", name = "txtA" }).Status);

            var refused = Send("POST", "/projects/p/modules/frm/controls", new { type = "Shell.Explorer.2", name = "web" });

            Assert.AreEqual(400, refused.Status);
            Assert.AreEqual(1, _store.Calls.Count(c => c == "Add"));
        }

        [TestCase("1bad")]
        [TestCase("has space")]
        [TestCase("")]
        public void Add_BadName_IsRefusedBeforeAskingTheUser(string name)
        {
            var response = Send("POST", "/projects/p/modules/frm/controls", new { type = "TextBox", name });

            Assert.AreEqual(400, response.Status);
            Assert.AreEqual(0, _asked);
            Assert.IsEmpty(_store.Calls);
        }

        [TestCase("Me")]
        [TestCase("Dim")]
        [TestCase("Print")]
        [TestCase("print")]
        [TestCase("WithEvents")]
        public void Add_ReservedWord_IsRefusedBeforeAskingTheUser(string name)
        {
            var response = Send("POST", "/projects/p/modules/frm/controls", new { type = "TextBox", name });

            Assert.AreEqual(400, response.Status);
            StringAssert.Contains("reserved", JObject.FromObject(response.Body)["error"]["message"].ToString());
            Assert.AreEqual(0, _asked);
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void Add_NameThatOnlyContainsAReservedWord_IsAccepted()
        {
            Assert.AreEqual(201, Send("POST", "/projects/p/modules/frm/controls", new { type = "TextBox", name = "txtPrint" }).Status);
        }

        [Test]
        public void Set_RenameToReservedWord_IsRefused()
        {
            Assert.AreEqual(400, Send("PUT", "/projects/p/modules/frm/controls/txt", new { properties = new { Name = "Dim" } }).Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Set_And_Delete_OfAControlThatAlreadyHasAReservedName_StillWork()
        {
            Assert.AreEqual(200, Send("PUT", "/projects/p/modules/frm/controls/Print", new { properties = new { Name = "txtPrint" } }).Status);
            Assert.AreEqual(200, Send("DELETE", "/projects/p/modules/frm/controls/Me").Status);
        }

        [Test]
        public void Add_NameInProperties_IsRefused()
        {
            var response = Send("POST", "/projects/p/modules/frm/controls", new { type = "TextBox", name = "txtA", properties = new { Name = "txtB" } });

            Assert.AreEqual(400, response.Status);
        }

        [Test]
        public void Set_WithoutProperties_IsRefused()
        {
            Assert.AreEqual(400, Send("PUT", "/projects/p/modules/frm/controls/txt", new { }).Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Set_RenameToBadName_IsRefused()
        {
            Assert.AreEqual(400, Send("PUT", "/projects/p/modules/frm/controls/txt", new { properties = new { Name = "9x" } }).Status);
        }

        [Test]
        public void Set_PassesTheProperties()
        {
            var response = Send("PUT", "/projects/p/modules/frm/controls/txt", new { caption = "Go", properties = new { Visible = false } });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Go", (string)_store.Properties["Caption"]);
            Assert.IsFalse((bool)_store.Properties["Visible"]);
        }

        [Test]
        public void Properties_MustBeAnObject_AndNamesMustLookLikeProperties()
        {
            Assert.AreEqual(400, Send("PUT", "/projects/p/modules/frm/controls/txt", new { properties = new[] { 1 } }).Status);
            Assert.AreEqual(400, Send("PUT", "/projects/p/modules/frm/controls/txt", new { properties = new Dictionary<string, object> { ["A b"] = 1 } }).Status);
        }

        [Test]
        public void Delete_AsksTheUser()
        {
            var response = Send("DELETE", "/projects/p/modules/frm/controls/txt");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(1, _asked);
            Assert.AreEqual("Delete", _store.Calls.Single());
        }

        [Test]
        public void Index_ListsTheRoutes()
        {
            var routes = JObject.FromObject(Send("GET", "").Body)["routes"].Select(r => (string)r).ToList();

            Assert.That(routes, Has.Some.Contains("/projects/{p}/modules/{m}/controls"));
            Assert.That(routes, Has.Some.Contains("/projects/{p}/modules/{m}/controls/{c}"));
        }
    }
}
