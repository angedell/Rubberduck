using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Api
{
    /// <summary>
    /// What agents can do with modules and forms. The REST routes and the MCP tools both come through here, so they cannot drift apart:
    /// the same checks, the same approval, the same errors.
    /// </summary>
    public class ModuleApi
    {
        //What the VBE accepts as the name of a module.
        private static readonly Regex ValidName = new Regex(@"^[A-Za-z][A-Za-z0-9_]{0,30}$", RegexOptions.Compiled);

        private readonly IModuleStore _store;
        private readonly ApiWriteApproval _approval;

        public ModuleApi(IModuleStore store, ApiWriteApproval approval)
        {
            _store = store;
            _approval = approval;
        }

        public IReadOnlyList<ProjectInfo> Projects() => _store.Projects();

        /// <summary>The modules, forms and classes of the project. Notes are not among them: they have their own tools (see <see cref="NotesApi"/>).</summary>
        public IReadOnlyList<ModuleInfo> Modules(string project) => _store.Modules(project).Where(module => !module.IsNote).ToList();

        public ModuleContent Read(string project, string module, bool includeText) => RejectNote(_store.Read(project, module, includeText));

        //A note is markdown behind comment marks. Through these tools an agent would have to keep the marks, the annotations and the line limit right by itself, so it is sent to the tools that do.
        private static T RejectNote<T>(T module) where T : ModuleInfo
        {
            if (module.IsNote)
            {
                throw new ApiException(409, "is_a_note",
                    $"'{module.Name}' is a note, not a code module. Use list_notes, read_note, write_note and delete_note (or the /notes routes) for it.");
            }

            return module;
        }

        private void RejectNote(string project, string module) => RejectNote(_store.Read(project, module, false));

        /// <summary>Like <see cref="Read(string, string, bool)"/>, and with the hidden attributes as data if asked for.</summary>
        public ModuleContent Read(string project, string module, bool includeText, bool includeAttributes)
        {
            var content = RejectNote(_store.Read(project, module, includeText || includeAttributes));
            if (includeAttributes)
            {
                content.Attributes = ModuleAttributes.Parse(content.Text);
                content.AttributesHash = ModuleAttributes.Hash(content.Attributes);
                if (!includeText)
                {
                    content.Text = null;
                }
            }

            return content;
        }

        /// <summary>Sets hidden attributes from data (see <see cref="ModuleAttributes"/>), so the caller never has to rebuild the exported header.</summary>
        /// <param name="expectedAttributesHash">If given, the attributes must still be the ones this hash (from a read with include=attributes) was made of; Hash alone does not notice an attribute that changed.</param>
        public ModuleInfo WriteAttributes(string client, string project, string module, Newtonsoft.Json.Linq.JObject attributes, string expectedHash, string expectedAttributesHash = null)
        {
            RequireHash(expectedHash);
            if (attributes == null)
            {
                throw ApiException.BadRequest("'attributes' must be an object.");
            }

            //Looks at the text first, so a bad name or value is refused before the user is asked anything.
            var current = RejectNote(_store.Read(project, module, true));
            if (!string.Equals(current.Hash, expectedHash, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.PreconditionFailed($"'{module}' changed since you read it. Read it again.");
            }

            if (expectedAttributesHash != null
                && !string.Equals(ModuleAttributes.Hash(ModuleAttributes.Parse(current.Text)), expectedAttributesHash, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.PreconditionFailed($"The attributes of '{module}' changed since you read them. Read it again with include=attributes.");
            }

            var text = ModuleAttributes.Apply(current.Text, attributes);
            _approval.Require(client);
            return _store.ReplaceText(project, module, text, expectedHash);
        }

        /// <summary>
        /// The editor drops the hidden attributes of a member (VB_UserMemId, VB_Description...) when the code that holds it is replaced; the module's own survive.
        /// Writes that change code put back the ones the new code still has a member for, and say so in <see cref="ModuleInfo.Notes"/>.
        /// Without this, replacing the code of a class loses its default member without a word.
        /// </summary>
        private ModuleInfo KeepingMemberAttributes(string project, string module, Func<ModuleInfo> write)
        {
            var before = ModuleAttributes.Parse(_store.Read(project, module, true).Text)["members"] as Newtonsoft.Json.Linq.JObject;
            var info = write();
            if (before == null || !before.HasValues)
            {
                return info;
            }

            var after = _store.Read(project, module, true);
            var present = ModuleAttributes.Parse(after.Text)["members"] as Newtonsoft.Json.Linq.JObject ?? new Newtonsoft.Json.Linq.JObject();
            var text = after.Text;
            var restored = new List<string>();
            foreach (var member in before.Properties())
            {
                var missing = new Newtonsoft.Json.Linq.JObject();
                foreach (var attribute in ((Newtonsoft.Json.Linq.JObject)member.Value).Properties())
                {
                    if ((present[member.Name] as Newtonsoft.Json.Linq.JObject)?[attribute.Name] == null)
                    {
                        missing[attribute.Name] = attribute.Value;
                    }
                }

                if (!missing.HasValues)
                {
                    continue;
                }

                try
                {
                    text = ModuleAttributes.Apply(text, new Newtonsoft.Json.Linq.JObject { ["members"] = new Newtonsoft.Json.Linq.JObject { [member.Name] = missing } });
                    restored.AddRange(missing.Properties().Select(a => member.Name + "." + a.Name));
                }
                catch (ApiException)
                {
                    //The new code has no such member any more; its attributes went with it, as they should.
                }
            }

            if (restored.Count == 0)
            {
                return info;
            }

            try
            {
                var result = _store.ReplaceText(project, module, text, after.Hash);
                result.Notes = new List<string> { "Put back the member attributes the editor dropped: " + string.Join(", ", restored) + "." };
                return result;
            }
            catch (ApiException exception)
            {
                info.Notes = new List<string> { $"The code was written, but these member attributes were dropped and could not be put back ({exception.Message}): {string.Join(", ", restored)}. Set them with 'attributes'." };
                return info;
            }
        }

        public ModuleInfo ReplaceCode(string client, string project, string module, string code, string expectedHash)
        {
            Require(code, "code");
            RequireHash(expectedHash);
            RejectHeaderLines(code);
            RejectNote(project, module);
            _approval.Require(client);
            return KeepingMemberAttributes(project, module, () => _store.ReplaceCode(project, module, code, expectedHash));
        }

        public ModuleInfo ReplaceText(string client, string project, string module, string text, string expectedHash)
        {
            Require(text, "text");
            RequireHash(expectedHash);
            RequireWholeModule(module, text);
            RejectNote(project, module);
            _approval.Require(client);
            return _store.ReplaceText(project, module, text, expectedHash);
        }

        public ModuleInfo EditLines(string client, string project, string module, int startLine, int endLine, string replacement, string expectedHash)
        {
            Require(replacement, "replacement");
            RequireHash(expectedHash);
            if (startLine < 1 || endLine < startLine - 1)
            {
                throw ApiException.BadRequest("startLine must be 1 or more, and endLine at least startLine - 1 (which inserts before startLine).");
            }

            RejectHeaderLines(replacement);
            RejectNote(project, module);
            _approval.Require(client);
            return KeepingMemberAttributes(project, module, () => _store.EditLines(project, module, startLine, endLine, replacement, expectedHash));
        }

        public ModuleInfo Create(string client, string project, string module, string kind, string code)
        {
            if (module == null || !ValidName.IsMatch(module))
            {
                throw ApiException.BadRequest("name must start with a letter and use only letters, digits and underscores (31 characters at most).");
            }

            kind = (kind ?? ModuleKinds.Standard).ToLowerInvariant();
            if (!ModuleKinds.Creatable.Contains(kind))
            {
                throw ApiException.BadRequest($"kind must be one of: {string.Join(", ", ModuleKinds.Creatable)}.");
            }

            RejectHeaderLines(code ?? string.Empty);
            _approval.Require(client);
            return _store.Create(project, module, kind, code ?? string.Empty);
        }

        public void Delete(string client, string project, string module, string expectedHash)
        {
            RequireHash(expectedHash);
            RejectNote(project, module);
            _approval.Require(client);
            _store.Delete(project, module, expectedHash);
        }

        public const int MaxBatchEdits = 50;

        /// <summary>
        /// Several changes at once, one per module. Everything that can be checked beforehand is (fields, hashes, names), so a stale hash
        /// in the last edit fails the batch before the first is applied. The changes themselves are applied in order; if one still fails,
        /// the error says which modules were already changed (their previous versions are in the backup folder).
        /// </summary>
        public BatchResult Batch(string client, string project, IReadOnlyList<ModuleEdit> edits)
        {
            if (edits == null || edits.Count == 0)
            {
                throw ApiException.BadRequest("'edits' must hold at least one edit.");
            }

            if (edits.Count > MaxBatchEdits)
            {
                throw ApiException.BadRequest($"At most {MaxBatchEdits} edits per batch.");
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var edit in edits)
            {
                if (edit == null || string.IsNullOrWhiteSpace(edit.Module))
                {
                    throw ApiException.BadRequest("Every edit needs a 'module'.");
                }

                if (!seen.Add(edit.Module))
                {
                    throw ApiException.BadRequest($"'{edit.Module}' is edited twice. Combine the changes into one edit per module.");
                }

                ValidateEdit(edit);
            }

            var existing = new HashSet<string>(_store.Modules(project).Select(m => m.Name), StringComparer.OrdinalIgnoreCase);
            foreach (var edit in edits)
            {
                if (Op(edit) == "create")
                {
                    if (existing.Contains(edit.Module))
                    {
                        throw ApiException.Conflict($"The project already has a module named '{edit.Module}'.");
                    }

                    continue;
                }

                var current = RejectNote(_store.Read(project, edit.Module, false));
                if (!string.Equals(current.Hash, edit.ExpectedHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw ApiException.PreconditionFailed($"'{edit.Module}' changed since you read it. Nothing was changed. Read it again and redo the batch.");
                }
            }

            _approval.Require(client);

            var result = new BatchResult();
            foreach (var edit in edits)
            {
                try
                {
                    result.Modules.Add(Apply(project, edit));
                }
                catch (ApiException failure)
                {
                    var done = result.Modules.Count == 0 ? "none" : string.Join(", ", result.Modules.Select(m => m.Name));
                    throw new ApiException(failure.Status, failure.Code,
                        $"The edit of '{edit.Module}' failed ({failure.Message}). Already changed: {done}. Nothing after '{edit.Module}' was applied.");
                }
            }

            return result;
        }

        private static string Op(ModuleEdit edit) => (edit.Op ?? string.Empty).Trim().ToLowerInvariant();

        private static void ValidateEdit(ModuleEdit edit)
        {
            switch (Op(edit))
            {
                case "replacecode":
                    Require(edit.Code, "code");
                    RequireHash(edit.ExpectedHash);
                    RejectHeaderLines(edit.Code);
                    break;
                case "replacetext":
                    Require(edit.Text, "text");
                    RequireHash(edit.ExpectedHash);
                    RequireWholeModule(edit.Module, edit.Text);
                    break;
                case "editlines":
                    Require(edit.Replacement, "replacement");
                    RequireHash(edit.ExpectedHash);
                    if (edit.StartLine < 1 || edit.EndLine < edit.StartLine - 1)
                    {
                        throw ApiException.BadRequest($"'{edit.Module}': startLine must be 1 or more, and endLine at least startLine - 1.");
                    }

                    RejectHeaderLines(edit.Replacement);
                    break;
                case "create":
                    if (!ValidName.IsMatch(edit.Module))
                    {
                        throw ApiException.BadRequest($"'{edit.Module}' is not a valid module name.");
                    }

                    if (!ModuleKinds.Creatable.Contains((edit.Kind ?? ModuleKinds.Standard).ToLowerInvariant()))
                    {
                        throw ApiException.BadRequest($"kind must be one of: {string.Join(", ", ModuleKinds.Creatable)}.");
                    }

                    RejectHeaderLines(edit.Code ?? string.Empty);
                    break;
                case "delete":
                    RequireHash(edit.ExpectedHash);
                    break;
                default:
                    throw ApiException.BadRequest($"'{edit.Module}': op must be one of: replaceCode, editLines, replaceText, create, delete.");
            }
        }

        private ModuleInfo Apply(string project, ModuleEdit edit)
        {
            switch (Op(edit))
            {
                case "replacecode":
                    return KeepingMemberAttributes(project, edit.Module, () => _store.ReplaceCode(project, edit.Module, edit.Code, edit.ExpectedHash));
                case "replacetext":
                    return _store.ReplaceText(project, edit.Module, edit.Text, edit.ExpectedHash);
                case "editlines":
                    return KeepingMemberAttributes(project, edit.Module, () => _store.EditLines(project, edit.Module, edit.StartLine, edit.EndLine, edit.Replacement, edit.ExpectedHash));
                case "create":
                    return _store.Create(project, edit.Module, (edit.Kind ?? ModuleKinds.Standard).ToLowerInvariant(), edit.Code ?? string.Empty);
                default:
                    _store.Delete(project, edit.Module, edit.ExpectedHash);
                    return new ModuleInfo { Project = project, Name = edit.Module, Kind = "deleted" };
            }
        }

        //'text' replaces the module as the VBE exports it. Text without the export's header is not that: the import would make a new module of it (named Module1, say) and drop the one meant.
        private static void RequireWholeModule(string module, string text)
        {
            var named = Regex.Match(text, @"^Attribute\s+VB_Name\s*=\s*""(?<name>[^""]*)""", RegexOptions.Multiline);
            if (!named.Success || !string.Equals(named.Groups["name"].Value, module, StringComparison.OrdinalIgnoreCase))
            {
                throw new ApiException(422, "text_not_a_module",
                    $"'text' has to be the whole module as exported, starting with its header and an 'Attribute VB_Name = \"{module}\"' line (read it with include=text). " +
                    "To change only the code, send 'code' (or use editLines); to change attributes, send 'attributes'. Nothing was changed.");
            }
        }

        private static void Require(string value, string field)
        {
            if (value == null)
            {
                throw ApiException.BadRequest($"'{field}' is required.");
            }
        }

        private static void RequireHash(string expectedHash)
        {
            if (string.IsNullOrWhiteSpace(expectedHash))
            {
                throw ApiException.PreconditionRequired("Send the hash from your last read as 'expectedHash' (or the If-Match header), so edits made in the meantime are not overwritten.");
            }
        }

        //Attribute lines are the VBE's hidden header; typed into the code they are a syntax error. Whole-module writes (text) are where they belong.
        private static void RejectHeaderLines(string code)
        {
            if (Regex.IsMatch(code, @"^\s*Attribute\s+VB_", RegexOptions.Multiline | RegexOptions.IgnoreCase))
            {
                throw new ApiException(422, "header_in_code",
                    "The code contains 'Attribute VB_...' lines. Send the code without its header, or replace the whole module with 'text' instead.");
            }
        }

        /// <summary>Hashes code the same way everywhere, ignoring the difference between line endings.</summary>
        public static string HashOf(string code)
        {
            var normalized = (code ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var hash = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(normalized));
                return BitConverter.ToString(hash, 0, 16).Replace("-", string.Empty).ToLowerInvariant();
            }
        }
    }
}
