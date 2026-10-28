using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IFormControlStore"/> on top of the VBE. The designer's controls are talked to late-bound, by property name, because
    /// what a control has depends on its type and the API takes the names as the caller writes them in VBA.
    /// </summary>
    public partial class VbeModuleStore
    {
        //What a listing tries on every control; a property the control does not have is left out.
        private static readonly string[] ListedProperties =
        {
            "Left", "Top", "Width", "Height", "TabIndex", "TabStop", "Visible", "Enabled", "Locked", "ControlTipText",
            "Caption", "Text", "Value", "ControlSource", "RowSource", "ColumnCount", "ColumnWidths", "BoundColumn", "Style", "ListStyle",
            "MatchRequired", "MultiSelect", "MultiLine", "WordWrap", "MaxLength", "PasswordChar", "ScrollBars", "TextAlign", "Min", "Max",
            "GroupName", "Default", "Cancel", "BackColor", "ForeColor"
        };

        public JObject ListControls(string project, string module)
            => OnUi(() => WithForm(project, module, (found, component) =>
            {
                var list = new JArray();
                using (var controls = component.Controls)
                {
                    foreach (var control in Raw(controls))
                    {
                        try
                        {
                            list.Add(Describe(control, NameOf(control), System.ComponentModel.TypeDescriptor.GetClassName(control)));
                        }
                        finally
                        {
                            Release(control);
                        }
                    }
                }

                return new JObject
                {
                    ["project"] = found.Name,
                    ["module"] = component.Name,
                    ["controls"] = list,
                    ["units"] = "Left, Top, Width and Height are in points, measured from the form's client area."
                };
            }));

        public JObject AddControl(string project, string module, string progId, string name, JObject properties)
            => OnUiWrite(() => WithForm(project, module, (found, component) =>
            {
                using (var controls = component.Controls)
                {
                    RequireNoControl(controls, name, module);
                    BackUpOnce(found, component);

                    object created;
                    try
                    {
                        created = Invoke(controls.Target, "Add", progId, name, true);
                    }
                    catch (Exception exception) when (IsComFailure(exception))
                    {
                        throw new ApiException(422, "control_refused", $"The designer refused to add a {progId} named '{name}': {Unwrap(exception).Message}");
                    }

                    try
                    {
                        Apply(created, name, properties);
                    }
                    catch (Exception)
                    {
                        //All or nothing: a control that was added but could not be set up is not left on the form.
                        try { Invoke(controls.Target, "Remove", name); } catch (Exception) { /* nothing more to do */ }
                        Release(created);
                        throw;
                    }

                    Release(created);
                    return Changed(found, component, controls, name);
                }
            }));

        public JObject SetControl(string project, string module, string name, JObject properties)
            => OnUiWrite(() => WithForm(project, module, (found, component) =>
            {
                using (var controls = component.Controls)
                {
                    var target = Find(controls, name, module);
                    BackUpOnce(found, component);
                    Apply(target, name, properties);
                    var newName = properties.Properties().FirstOrDefault(p => p.Name.Equals("Name", StringComparison.OrdinalIgnoreCase))?.Value.Value<string>() ?? name;
                    Release(target);
                    return Changed(found, component, controls, newName);
                }
            }));

        public void DeleteControl(string project, string module, string name)
            => OnUiWrite(() => WithForm(project, module, (found, component) =>
            {
                using (var controls = component.Controls)
                {
                    Release(Find(controls, name, module));
                    BackUpOnce(found, component);
                    try
                    {
                        Invoke(controls.Target, "Remove", name);
                    }
                    catch (Exception exception) when (IsComFailure(exception))
                    {
                        throw new ApiException(422, "control_refused", $"The designer refused to delete '{name}': {Unwrap(exception).Message}");
                    }

                    _parseManager.MarkAsModified(new QualifiedModuleName(component));
                    return true;
                }
            }));

        private T WithForm<T>(string project, string module, Func<IVBProject, IVBComponent, T> work)
        {
            var found = FindProject(project);
            using (var components = found.VBComponents)
            using (var component = Find(components, module))
            {
                if (component.Type != ComponentType.UserForm)
                {
                    throw new ApiException(422, "not_a_form", $"'{module}' is a {KindOf(component.Type)} module, not a form, so it has no controls.");
                }

                return work(found, component);
            }
        }

        private JObject Changed(IVBProject project, IVBComponent component, IControls controls, string name)
        {
            //The parser declares a form's controls as members of the form.
            _parseManager.MarkAsModified(new QualifiedModuleName(component));
            var target = Find(controls, name, component.Name);
            try
            {
                var control = Describe(target, name, System.ComponentModel.TypeDescriptor.GetClassName(target));
                control["project"] = project.Name;
                control["module"] = component.Name;
                return control;
            }
            finally
            {
                Release(target);
            }
        }

        //The wrapper's enumeration soft-casts to Forms.Control, which an Image is not, so it comes out as a null reference and would be skipped.
        //The caller owns each object and releases it.
        private static IEnumerable<object> Raw(IControls controls)
            => controls.IsWrappingNullReference ? Enumerable.Empty<object>() : ((System.Collections.IEnumerable)controls.Target).Cast<object>();

        private static string NameOf(object control) => Get(control, "Name") as string;

        private static JObject Describe(object target, string name, string className)
        {
            var result = new JObject
            {
                ["name"] = name,
                ["type"] = FriendlyType(className)
            };

            foreach (var property in ListedProperties)
            {
                if (property == "Value" && result["text"] != null)
                {
                    continue;
                }

                try
                {
                    var value = Get(target, property);
                    if (value == null || value is string || value is bool || value is IConvertible)
                    {
                        result[char.ToLowerInvariant(property[0]) + property.Substring(1)] = value == null ? null : JToken.FromObject(value);
                    }
                }
                catch (Exception exception) when (IsComFailure(exception))
                {
                    //The control does not have this property.
                }
            }

            return result;
        }

        //TypeDescriptor says "Forms.TextBox.1" or the like; the caller knows it as TextBox.
        private static string FriendlyType(string className)
        {
            var name = className ?? "Control";
            if (name.StartsWith("Forms.", StringComparison.OrdinalIgnoreCase)) name = name.Substring("Forms.".Length);
            if (name.StartsWith("MSForms.", StringComparison.OrdinalIgnoreCase)) name = name.Substring("MSForms.".Length);
            if (name.EndsWith(".1", StringComparison.Ordinal)) name = name.Substring(0, name.Length - 2);
            return name;
        }

        private static void RequireNoControl(IControls controls, string name, string module)
        {
            foreach (var control in Raw(controls))
            {
                try
                {
                    if (string.Equals(NameOf(control), name, StringComparison.OrdinalIgnoreCase))
                    {
                        throw ApiException.Conflict($"'{module}' already has a control named '{name}'.");
                    }
                }
                finally
                {
                    Release(control);
                }
            }
        }

        //The caller owns the returned COM object and releases it.
        private static object Find(IControls controls, string name, string module)
        {
            foreach (var control in Raw(controls))
            {
                try
                {
                    var found = NameOf(control);
                    if (string.Equals(found, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return Invoke(controls.Target, "Item", found);
                    }
                }
                finally
                {
                    Release(control);
                }
            }

            throw ApiException.NotFound($"'{module}' has no control '{name}'. List them with list_controls.");
        }

        private static void Apply(object target, string name, JObject properties)
        {
            //All or nothing: what was set before a refusal is put back, so a failed request leaves the control as it was.
            var applied = new List<KeyValuePair<string, object>>();
            foreach (var property in properties.Properties())
            {
                var value = ToCom(property);
                object before = null;
                var known = false;
                try
                {
                    try { before = Get(target, property.Name); known = true; } catch (Exception getter) when (IsComFailure(getter)) { /* write-only or unknown; the set below reports it */ }
                    Set(target, property.Name, value);
                    if (known)
                    {
                        applied.Add(new KeyValuePair<string, object>(property.Name, before));
                    }
                }
                catch (Exception exception) when (IsComFailure(exception))
                {
                    for (var i = applied.Count - 1; i >= 0; i--)
                    {
                        try { Set(target, applied[i].Key, applied[i].Value); } catch (Exception) { /* nothing more to do */ }
                    }

                    throw new ApiException(422, "property_refused", $"'{name}' refused {property.Name} = {property.Value}: {Unwrap(exception).Message}");
                }
            }
        }

        private static object ToCom(JProperty property)
        {
            switch (property.Value.Type)
            {
                case JTokenType.Boolean:
                    return property.Value.Value<bool>();
                case JTokenType.String:
                    return property.Value.Value<string>();
                case JTokenType.Integer:
                    var whole = property.Value.Value<long>();
                    return whole >= int.MinValue && whole <= int.MaxValue ? (object)(int)whole : whole;
                case JTokenType.Float:
                    return property.Value.Value<double>();
                default:
                    throw ApiException.BadRequest($"'{property.Name}' must be a string, number or true/false.");
            }
        }

        private static object Get(object target, string property)
        {
            try
            {
                return target.GetType().InvokeMember(property, BindingFlags.GetProperty, null, target, null);
            }
            catch (TargetInvocationException exception)
            {
                throw exception.InnerException ?? exception;
            }
        }

        private static void Set(object target, string property, object value)
        {
            try
            {
                target.GetType().InvokeMember(property, BindingFlags.SetProperty, null, target, new[] { value });
            }
            catch (TargetInvocationException exception)
            {
                throw exception.InnerException ?? exception;
            }
        }

        private static object Invoke(object target, string method, params object[] arguments)
        {
            try
            {
                return target.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, target, arguments);
            }
            catch (TargetInvocationException exception)
            {
                throw exception.InnerException ?? exception;
            }
        }

        private static bool IsComFailure(Exception exception)
        {
            var inner = Unwrap(exception);
            return inner is COMException || inner is MissingMemberException || inner is InvalidCastException || inner is ArgumentException || inner is FormatException;
        }

        private static Exception Unwrap(Exception exception) => (exception as TargetInvocationException)?.InnerException ?? exception;

        private static void Release(object target)
        {
            if (target != null && Marshal.IsComObject(target))
            {
                Marshal.ReleaseComObject(target);
            }
        }
    }
}
