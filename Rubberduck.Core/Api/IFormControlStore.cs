using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// The controls on a UserForm's designer, as the API sees them. Implementations throw <see cref="ApiException"/> for what the caller got wrong.
    /// Positions and sizes are in points, like the designer's Properties window shows them.
    /// </summary>
    public interface IFormControlStore
    {
        /// <summary>The form's controls, each with its name, type and the properties it has among the ones that matter for layout and behavior.</summary>
        JObject ListControls(string project, string module);

        /// <summary>Adds a control of the given ProgID (see <see cref="FormApi.ControlTypes"/>) and sets <paramref name="properties"/> on it. All or nothing: if a property is refused the control is not left behind.</summary>
        JObject AddControl(string project, string module, string progId, string name, JObject properties);

        /// <summary>Sets properties of an existing control. 'Name' renames it.</summary>
        JObject SetControl(string project, string module, string name, JObject properties);

        void DeleteControl(string project, string module, string name);
    }
}
