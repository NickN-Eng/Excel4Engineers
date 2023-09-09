using System;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute for creating a button where the label text can be bound to a DynamicText and change dynamically
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public class DynamicButtonAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// An attribute for creating a button where the label text can be bound to a DynamicText and change dynamically
        /// </summary>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="dynamicLabelProperty">The property name of the DynamicText property which this label text is bound to</param>
        /// <param name="isLarge">True if button is large</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        /// <param name="iconImage">Key to the Icon image within Resources</param>
        public DynamicButtonAttribute(string groupText, string dynamicLabelProperty, bool isLarge = false, string tooltipHeader = null, string tooltipDescription = null, string iconImage = null) : base(/*label*/"", groupText, tooltipHeader, tooltipDescription, iconImage)
        {
            IsLarge = isLarge;
            DynamicLabelProperty = dynamicLabelProperty;
        }

        /// <summary>
        /// The property name of the DynamicText property which this label text is bound to
        /// </summary>
        public string DynamicLabelProperty { get; set; }

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
            var dynamicLabel = ribbon.FindProperty<DynamicText>(DynamicLabelProperty);
            dynamicLabel.Initialise(ribbon, id);
            ribbon.RegisterDynamicLabel(id, dynamicLabel);
            RegisterImage(id, ribbon);


            return $@"<button id='{id}' getLabel='GetDynamicLabel' onAction='{CallerName}' {(IsLarge ? @" size = 'large'" : "")}   {MiscTextXml()}/>";
        }
    }

}
