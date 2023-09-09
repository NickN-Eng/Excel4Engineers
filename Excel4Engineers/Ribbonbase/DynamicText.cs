namespace Excel4Engineers
{

    /// <summary>
    /// Container for a text field which is used in the Ribbon UI. 
    /// Notifies the RibbonUI for refresh when the Text is updated
    /// </summary>
    public class DynamicText
    {
        private string _Text;
        /// <summary>
        /// The dynamic text value that can be bound to a label text or combobox value
        /// </summary>
        public string Text
        {
            get => _Text; 
            set
            {
                _Text = value;
                Ribbon.RibbonUI.InvalidateControl(Id);
            }
        }

        /// <summary>
        /// The id of the ribbon UI element
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// The reference to the Ribbon UI
        /// </summary>
        RibbonControllerBase Ribbon { get; set; }

        /// <summary>
        /// Create a DynamicText, a container for a text field which is used in the Ribbon UI
        /// and can notify the RibbonUI for refresh when the Text is updated
        /// </summary>
        public DynamicText(string label)
        {
            _Text = label;
        }

        /// <summary>
        /// This dynamic object must be initalised by the Ribbon and element id info to function dynamically
        /// </summary>
        public void Initialise(RibbonControllerBase ribbon, string id)
        {
            Ribbon = ribbon;
            Id = id;

        }
    }

}
