using System;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute for creating a button WHICH IS DISABLED where the label text can be bound to a DynamicText and change dynamically
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class DynamicButtonLabelAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// Base attribute for ribbon ui elements
        /// </summary>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="isLarge">True if button is large</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        /// <param name="iconImage">Key to the Icon image within Resources</param>
        public DynamicButtonLabelAttribute(string groupText, bool isLarge = false, string tooltipHeader = null, string tooltipDescription = null, string iconImage = null) : base("", groupText, tooltipHeader, tooltipDescription, iconImage)
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
            string id = $"label{ribbon.GetNextUIElementIndex()}";
            var dynamicLabel = ribbon.FindProperty<DynamicText>(CallerName);
            dynamicLabel.Initialise(ribbon, id);
            ribbon.RegisterDynamicLabel(id, dynamicLabel);
            RegisterImage(id, ribbon);

            return $@"<button id='{id}' getLabel='GetDynamicLabel' enabled='false' {(IsLarge ? @" size = 'large'" : "")}  {MiscTextXml()}/>";
        }
    }

}
