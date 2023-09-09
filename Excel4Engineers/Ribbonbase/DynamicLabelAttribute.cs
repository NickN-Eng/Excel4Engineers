using System;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute for creating a ribbon UI label where the label text can be bound to a DynamicText and change dynamically
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class DynamicLabelAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// An attribute for creating a ribbon UI label where the label text can be bound to a DynamicText and change dynamically
        /// </summary>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        public DynamicLabelAttribute(string groupText, string tooltipHeader = null, string tooltipDescription = null) : base("", groupText, tooltipHeader, tooltipDescription)
        {
        }

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


            return $@"<labelControl id='{id}' getLabel='GetDynamicLabel' {MiscTextXml()}/>";
        }
    }

}
