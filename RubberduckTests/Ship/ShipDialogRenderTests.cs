using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;
using Rubberduck.Ship;
using Rubberduck.UI.CodeExplorer.Ship;

namespace RubberduckTests.Ship
{
    /// <summary>Draws the dialog off screen, so layout can be looked at without Excel. The PNGs go to the temp folder; nothing is asserted about pixels.</summary>
    [TestFixture]
    [Category("Ship")]
    [Apartment(System.Threading.ApartmentState.STA)]
    public class ShipDialogRenderTests
    {
        private static string Save(Form form, string name)
        {
            //A form that was never shown draws an empty client area, so it is shown far off screen first.
            form.StartPosition = FormStartPosition.Manual;
            form.ShowInTaskbar = false;
            form.Location = new Point(-32000, -32000);
            form.Height = Math.Max(form.Height, 1500); //tall enough to show the whole settings page without scrolling
            if (!form.Visible)
            {
                form.Show();
            }

            form.PerformLayout();
            Application.DoEvents();
            var bitmap = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
            var path = Path.Combine(Path.GetTempPath(), "ship-dialog-" + name + ".png");
            bitmap.Save(path, ImageFormat.Png);
            return path;
        }

        /// <summary>The form cannot be taller than the screen, so the part of the settings page below the fold is drawn by scrolling it to its end.</summary>
        private static void ScrollSettingsToEnd(Form form)
        {
            foreach (var panel in form.Controls.OfType<Panel>().Where(p => p.Visible && p.AutoScroll))
            {
                panel.AutoScrollPosition = new Point(0, int.MaxValue / 4);
            }

            Application.DoEvents();
        }

        [Test]
        public async Task RendersEveryPage()
        {
            var report = new ShipReport { DryRun = true };
            var notes = new ShipStepReport { Step = ShipSteps.Notes, Removed = 2 };
            notes.Details.Add("Notes1");
            notes.Details.Add("Notes2");
            report.Steps.Add(notes);
            report.Steps.Add(new ShipStepReport { Step = ShipSteps.Comments, Removed = 41 });
            report.Unparsed.Add("Module3: line 12 unexpected token");
            report.Warnings.Add("Pending resource changes were not shipped");
            var service = new ShipViewModelTests.FakeShipService
            {
                Settings = new ShipSettings { OutputFolder = @"..\dist" },
                Resolved = Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar),
                OnPreview = s => report,
                OnShip = (s, e, p) => { p.Report("Copying the workbook"); return new ShipReport { OutputPath = Path.Combine(Path.GetTempPath(), "Book.xlsm") }; }
            };
            var model = new ShipViewModel(service, null, "Book", new ShipViewModelTests.FakeProbe());
            model.Load();
            await model.RefreshPreviewAsync();

            using (var dialog = new ShipDialog(model, null))
            {
                dialog.Render();
                TestContext.WriteLine(Save(dialog, "1-settings"));

                await model.ProceedAsync();
                dialog.Render();
                TestContext.WriteLine(Save(dialog, "2-confirm"));

                await model.ConfirmAsync();
                dialog.Render();
                TestContext.WriteLine(Save(dialog, "3-done"));
            }

            Assert.AreEqual(ShipPage.Done, model.Page);
        }

        [Test]
        public async Task RendersTreeShakingWithRemovalsAndDynamicCalls()
        {
            var report = new ShipReport { DryRun = true };
            var shaking = new ShipStepReport { Step = ShipSteps.TreeShaking, Removed = 3 };
            shaking.Details.Add("Module1.OldHelper (procedure)");
            shaking.Details.Add("Module1.mLegacyCount (variable)");
            shaking.Details.Add("Utilities.FormatLegacy (procedure)");
            report.Steps.Add(shaking);
            report.DynamicCalls.Add("Module2.RunByName line 14: Application.Run with a computed name");
            report.DynamicCalls.Add("Module2.Schedule line 31: Application.OnTime with a computed name");
            var settings = new ShipSettings { OutputFolder = @"..\dist" };
            settings.Steps[ShipSteps.TreeShaking] = true;
            settings.TreeShaking.Mode = ShipTreeShakingMode.EntryPointsOnly;
            settings.TreeShaking.Keep.Add("Module2.Report");
            settings.TreeShaking.Keep.Add("Refresh");
            var service = new ShipViewModelTests.FakeShipService
            {
                Settings = settings,
                Resolved = Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar),
                OnPreview = s => report
            };
            var model = new ShipViewModel(service, null, "Book", new ShipViewModelTests.FakeProbe());
            model.Load();
            await model.RefreshPreviewAsync();

            using (var dialog = new ShipDialog(model, null))
            {
                dialog.Render();
                TestContext.WriteLine(Save(dialog, "4-treeshaking-settings"));
                ScrollSettingsToEnd(dialog);
                TestContext.WriteLine(Save(dialog, "4b-treeshaking-settings-bottom"));

                await model.ProceedAsync();
                dialog.Render();
                TestContext.WriteLine(Save(dialog, "5-treeshaking-confirm"));
            }

            Assert.AreEqual(ShipPage.Confirm, model.Page);
        }
    }
}
