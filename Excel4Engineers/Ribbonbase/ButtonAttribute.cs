using System;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute to be attatched to method within the RibbonController which turns it into a UI button
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public class ButtonAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// An attribute to be attatched to method within the RibbonController which turns it into a UI button
        /// </summary>
        /// <param name="label">Labal text</param>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="isLarge">True if button is large</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        /// <param name="iconImage">Key to the Icon image within Resources</param>
        public ButtonAttribute(string label, string groupText, bool isLarge = false, string tooltipHeader = null, string tooltipDescription = null, string iconImage = null) : base(label, groupText, tooltipHeader, tooltipDescription, iconImage)
        {
            IsLarge = isLarge;
        }

        /// <summary>
        /// If true, this button is large
        /// </summary>
        public bool IsLarge { get; set; } = false;

        /// <summary>
        /// Creates the xml and registers all dynamic properties with the ribbon control
        /// </summary>
        public override string InitAndGetXml(RibbonControllerBase ribbon)
        {
            string id = $"button{ribbon.GetNextUIElementIndex()}";
            RegisterImage(id, ribbon);

            return $@"<button id='{id}' label='{Label}' onAction='{CallerName}' {(IsLarge ? @" size = 'large'" : "")} {MiscTextXml()}/>";
        }
    }

}
