using System;
using System.Text;

namespace Excel4Engineers
{
    /// <summary>
    /// An attribute to be attatched to a DynamicText property within the RibbonController which turns it into a UI combobox
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ComboboxAttribute : RibbonElementAttribute
    {
        /// <summary>
        /// An attribute to be attatched to a DynamicText property within the RibbonController which turns it into a UI combobox
        /// </summary>
        /// <param name="label">Labal text</param>
        /// <param name="groupText">Group for this UI element</param>
        /// <param name="options">Static list of text options for the combobox</param>
        /// <param name="tooltipHeader">Tooltip header text shown when user hovers over</param>
        /// <param name="tooltipDescription">Tooltip description text shown when user hovers over</param>
        public ComboboxAttribute(string label, string groupText, string[] options, string tooltipHeader = null, string tooltipDescription = null) : base(label, groupText, tooltipHeader, tooltipDescription)
        {
            Options = options;
        }

        /// <summary>
        /// Static list of text options for the combobox
        /// </summary>
        public string[] Options { get; set; }

        /// <summary>
        /// Creates the xml and registers all dynamic properties with the ribbon control
        /// </summary>
        public override string InitAndGetXml(RibbonControllerBase ribbon)
        {
            string id = $"comboBox{ribbon.GetNextUIElementIndex()}";
            var dynamicLabel = ribbon.FindProperty<DynamicText>(CallerName);
            dynamicLabel.Initialise(ribbon, id);
            ribbon.RegisterDynamicDynamicComboboxText(id, dynamicLabel);
            RegisterImage(id, ribbon);

            var sb = new StringBuilder();
            sb.AppendLine($@"<comboBox id ='{id}' {LabelTextXml()} getText='GetDynamicComboboxText' onChange='ComboboxOnChange' {MiscTextXml()}>");
            foreach (var opt in Options)
            {
                sb.AppendLine($@"<item id ='comboItem{ribbon.GetNextUIElementIndex()}' label='{opt}'/>");
            }
            sb.AppendLine("</comboBox>");

            return sb.ToString();
        }
    }

}
