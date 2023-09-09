using System;
using System.Drawing;
using Excel4Engineers.Properties;

namespace Excel4Engineers
{
    [AttributeUsage(AttributeTargets.Method)]
    public abstract class RibbonElementAttribute : Attribute
    {
        /// <summary>
        /// Base attribute for ribbon ui elements
        /// </summary>
        /// <param name="label">Labal text</param>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        /// <param name="iconImage">Key to the Icon image within Resources</param>
        public RibbonElementAttribute(string label, string groupText, string tooltipHeader = null, string tooltipDescription = null, string iconImage = null)
        {
            Label = label;
            GroupText = groupText;
            TooltipHeader = tooltipHeader;
            TooltipDescription = tooltipDescription;
            IconImage = iconImage;
        }

        /// <summary>
        /// The label of the ribbon UI element
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// The group text. All ribbon UI element with the same group will be placed into a ribbon UI group.
        /// </summary>
        public string GroupText { get; set; }

        /// <summary>
        /// The name of the property/method on which this attribute is attatched.
        /// This is populated by the RibbonControllerBase through reflection and used during InitAndGetXml
        /// </summary>
        public string CallerName { get; set; }

        /// <summary>
        /// The tooltip header text
        /// </summary>
        public string TooltipHeader { get; set; }

        /// <summary>
        /// The tooltip description
        /// </summary>
        public string TooltipDescription { get; set; }

        /// <summary>
        /// Key to the Icon image within Resources
        /// </summary>
        public string IconImage { get; set; }

        /// <summary>
        /// Creates the xml text for Tooltip headers and images.
        /// Should be called within the InitAndGetXml method of inherited members
        /// </summary>
        protected string MiscTextXml() => (TooltipHeader != null ? $" screentip='{TooltipHeader}' " : "") + (TooltipDescription != null ? $" supertip='{TooltipDescription}' " : "") + (IconImage != null ? $"  getImage='GetImage' " : "");

        /// <summary>
        /// Creates the xml text which accounts for the Label being null, in which case the label will be hidden
        /// Should be called within the InitAndGetXml method of inherited members
        /// </summary>
        /// <returns></returns>
        public string LabelTextXml() => Label == null ? " showLabel='false' " : $" label='{Label}'";

        /// <summary>
        /// Adds the relevant icon image to the ribbon icon library, provided this control has an icon.
        /// Should be called within the InitAndGetXml method of inherited members
        /// </summary>
        protected void RegisterImage(string id, RibbonControllerBase ribbon)
        {
            if (IconImage != null)
            {
                var image = (Image)Resources.ResourceManager.GetObject(IconImage);
                ribbon.RegisterIconImage(id, image);
            }
        }

        /// <summary>
        /// Creates the xml and registers all dynamic properties with the ribbon control
        /// </summary>
        public abstract string InitAndGetXml(RibbonControllerBase ribbon);
    }

}
