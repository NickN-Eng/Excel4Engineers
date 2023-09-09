namespace Excel4Engineers
{
    /// <summary>
    /// Container for a bool field which is used in the Ribbon UI. 
    /// Notifies the RibbonUI for refresh when the Text is updated
    /// </summary>
    public class DynamicBool
    {
        private bool _Value;
        /// <summary>
        /// The boolean value that can used to bind to a checkbox/enable property
        /// </summary>
        public bool Value
        {
            get => _Value;
            set
            {
                _Value = value;
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
        public DynamicBool(bool value)
        {
            _Value = value;
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
