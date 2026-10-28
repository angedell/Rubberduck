using System.Collections.Generic;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Moq;

namespace RubberduckTests.Ship
{
    /// <summary>The parser Rubberduck itself uses for a module's text (with the preprocessor, so #If branches are evaluated), without a VBE.</summary>
    internal static class ShipTestParser
    {
        public static IStringParser Create(Dictionary<string, short> compilationConstants = null)
        {
            var arguments = new Mock<ICompilationArgumentsProvider>();
            var predefined = new VBAPredefinedCompilationConstants(7.1);
            arguments.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>())).Returns(compilationConstants ?? new Dictionary<string, short>());
            arguments.Setup(m => m.PredefinedCompilationConstants).Returns(() => predefined);

            var preprocessorErrors = new PreprocessingParseErrorListenerFactory();
            //The preprocessor reads the project's #Const values from the cache, which has to be told to load them.
            var cache = new CompilationArgumentsCache(arguments.Object);
            cache.ReloadCompilationArguments(new[] { "project-1" });
            var preprocessor = new VBAPreprocessor(new VBAPreprocessorParser(preprocessorErrors, preprocessorErrors), cache);
            var mainErrors = new MainParseErrorListenerFactory();
            return new TokenStreamParserStringParserAdapterWithPreprocessing(new SimpleVBAModuleTokenStreamProvider(), new VBATokenStreamParser(mainErrors, mainErrors), preprocessor);
        }

        public static string Lines(params string[] lines) => string.Join("\r\n", lines) + "\r\n";
    }
}
