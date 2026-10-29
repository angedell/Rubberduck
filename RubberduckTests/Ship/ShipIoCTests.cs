using System;
using System.Collections.Generic;
using System.Reflection;
using Castle.MicroKernel.Registration;
using Castle.Windsor;
using Moq;
using NUnit.Framework;
using Rubberduck.Root;
using Rubberduck.Runtime;
using Rubberduck.Settings;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.VbeRuntime;
using RubberduckTests.Mocks;

namespace RubberduckTests.Ship
{
    /// <summary>How Windsor builds Ship: the container must hand the service the parser-state analysis, and a container without one must still build the service.</summary>
    [TestFixture]
    [Category("Ship")]
    [Category("IoC_Registration")]
    public class ShipIoCTests
    {
        private static object Field(object instance, string name)
        {
            var field = instance.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
            return field.GetValue(instance);
        }

        [Test]
        public void TheInstallerGivesTheServiceTheAnalysisOverTheParserState()
        {
            var ideMock = new MockVbeBuilder().Build();
            var sourceFileHandler = new Mock<ITempSourceFileHandler>().Object;
            ideMock.Setup(m => m.TempSourceFileHandler).Returns(() => sourceFileHandler);
            var addin = new MockAddInBuilder().Build().Object;

            IWindsorContainer container = null;
            try
            {
                try
                {
                    container = new WindsorContainer().Install(new RubberduckIoCInstaller(ideMock.Object, addin, new GeneralSettings { EnableExperimentalFeatures = new List<ExperimentalFeature> { new ExperimentalFeature() } },
                        new Mock<IVbeNativeApi>().Object, new Mock<IBeepInterceptor>().Object));
                }
                catch (Exception exception)
                {
                    Assert.Inconclusive($"Unable to register. {Environment.NewLine} {exception}");
                }

                var analysis = container.Resolve<IShipAnalysis>();
                var service = container.Resolve<IShipService>();
                var pipeline = Field(service, "_pipeline");

                Assert.IsInstanceOf<ParserStateShipAnalysis>(analysis);
                Assert.AreSame(analysis, Field(pipeline, "_analysis"), "The service has to be built with the registered analysis.");
            }
            finally
            {
                container?.Dispose();
            }
        }

        [Test]
        public void AContainerWithoutAnAnalysisStillBuildsTheServiceWithTreeShakingUnavailable()
        {
            using (var container = new WindsorContainer())
            {
                container.Register(Component.For<System.IO.Abstractions.IFileSystem>().Instance(new System.IO.Abstractions.FileSystem()));
                container.Register(Component.For<Rubberduck.PackageResources.IPackageResourceService>().Instance(new Mock<Rubberduck.PackageResources.IPackageResourceService>().Object));
                container.Register(Component.For<Rubberduck.Parsing.VBA.Parsing.IStringParser>().Instance(ShipTestParser.Create()));
                container.Register(Component.For<IShipService>().ImplementedBy<ShipService>());

                var service = container.Resolve<IShipService>();

                Assert.IsNull(Field(Field(service, "_pipeline"), "_analysis"));
            }
        }
    }
}
