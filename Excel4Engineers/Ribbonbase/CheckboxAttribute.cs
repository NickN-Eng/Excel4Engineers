using System;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute to be attatched to a DynamicBool property within the RibbonController which turns it into a UI checkbox
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class CheckboxAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// An attribute to be attatched to a DynamicBool property within the RibbonController which turns it into a UI checkbox
        /// </summary>
        /// <param name="label">Label text</param>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        public CheckboxAttribute(string label, string groupText, string tooltipHeader = null, string tooltipDescription = null) : base(label, groupText, tooltipHeader, tooltipDescription)
        {
        }

        /// <summary>
        /// Creates the xml and registers all dynamic properties with the ribbon control
        /// </summary>
        public override string InitAndGetXml(RibbonControllerBase ribbon)
        {
            string id = $"checkbox{ribbon.GetNextUIElementIndex()}";
            var dynamicBool = ribbon.FindProperty<DynamicBool>(CallerName);
            dynamicBool.Initialise(ribbon, id);
            ribbon.RegisterCheckboxValue(id, dynamicBool);
            RegisterImage(id, ribbon);

            return $@"<checkBox id='{id}' label='{Label}' getPressed='CheckboxGetPressed' onAction='CheckboxOnAction' {MiscTextXml()}/>";
        }
    }

}
